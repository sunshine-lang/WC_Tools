"""Source-language TM retrieval: cached embeddings + BM25 recall + remote reranking.

Only source text is retrieved. Original bilingual strings remain in TMMatcher;
Context and terminology are left for the downstream translator to interpret.
"""
import hashlib
import json
import math
import os
import re
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import urlsplit
from zipfile import BadZipFile

import numpy as np
import requests
from rank_bm25 import BM25Okapi


from .language import language_pair, source_tokens


SEMANTIC_POLICY = 'wc-pair-semantic-v2'
ROOT = Path(__file__).resolve().parents[1]


def word_tokens(text, chinese=False):
    return source_tokens(text, chinese)


class SemanticTMRetriever:
    def __init__(self, config_path=None, cache_dir=None, language=None):
        self.language = language_pair(language)
        try:
            cfg = json.loads(Path(config_path or ROOT / 'config/config.json').read_text('utf-8'))
            self.embedding = cfg.get('embedding', {})
            self.rerank_config = cfg.get('rerank', {})
            for label, section in (('向量', self.embedding), ('重排', self.rerank_config)):
                endpoint = urlsplit(section.get('api_url', ''))
                if not (section.get('api_key') and section.get('model_name')
                        and endpoint.scheme in ('https', 'http') and endpoint.hostname):
                    raise ValueError(f'TM {label}配置不完整，请检查 config/config.json')
            self.batch_size = max(1, min(128, int(self.embedding.get('batch_size', 10))))
            self.embed_workers = max(1, min(8, int(self.embedding.get('max_workers', 5))))
            self.rank_workers = max(1, min(8, int(self.rerank_config.get('max_workers', 5))))
            self.embed_timeout = max(1, float(self.embedding.get('timeout', 60)))
            self.rank_timeout = max(1, float(self.rerank_config.get('timeout', 30)))
            self.recall_k = max(1, min(100, int(self.rerank_config.get('recall_k', 30))))
            self.final_k = max(1, min(self.recall_k, int(self.rerank_config.get('final_k', 5))))
            self.threshold = float(self.rerank_config.get('threshold', 0.65))
            self.retries = max(1, min(3, int(self.rerank_config.get('request_retries', 3))))
            self.backoff = max(0.1, float(self.rerank_config.get('retry_backoff', 2)))
            if not all(math.isfinite(v) for v in (self.threshold, self.backoff,
                                                   self.embed_timeout, self.rank_timeout)):
                raise ValueError('TM 模型配置中的数值必须有限')
        except (OSError, TypeError, AttributeError, json.JSONDecodeError) as error:
            raise ValueError('无法读取 TM 模型配置，请检查 config/config.json 的 embedding / rerank') from error
        self.cache_dir = Path(cache_dir or ROOT / 'workspace/tm_cache')
        self._local = threading.local()
        self.vectors = None
        self.cache_hit = False

    def status(self):
        return {'ready': True, 'embedding_model': self.embedding['model_name'],
                'rerank_model': self.rerank_config['model_name'], 'recall_k': self.recall_k,
                'final_k': self.final_k, 'threshold': self.threshold, 'cache_hit': self.cache_hit}

    def _post(self, section, payload, timeout, label):
        if not hasattr(self._local, 'session'):
            self._local.session = requests.Session()
        for attempt in range(self.retries):
            try:
                response = self._local.session.post(
                    section['api_url'], json=payload,
                    headers={'Authorization': f"Bearer {section['api_key']}", 'Content-Type': 'application/json'},
                    timeout=timeout, allow_redirects=False)
            except (requests.Timeout, requests.ConnectionError):
                failure = '请求超时或连接失败'
            except requests.RequestException as error:
                raise ValueError(f'TM {label}请求失败，请检查接口配置') from error
            else:
                try:
                    if response.status_code == 200:
                        try:
                            return response.json()
                        except ValueError as error:
                            raise ValueError(f'TM {label}返回的 JSON 无效') from error
                    failure = f'接口返回 HTTP {response.status_code}'
                    if response.status_code not in (429, 500, 502, 503, 504):
                        raise ValueError(f'TM {label}{failure}，请检查配置')
                finally:
                    response.close()
            if attempt + 1 == self.retries:
                raise ValueError(f'TM {label}{failure}（已尝试 {self.retries} 次），请重试本步')
            time.sleep(self.backoff * (2 ** attempt))

    @staticmethod
    def _vectors(value, count):
        try:
            matrix = np.asarray(value, dtype=np.float32)
            if matrix.ndim != 2 or matrix.shape[0] != count or not matrix.shape[1]:
                raise ValueError()
            norms = np.linalg.norm(matrix.astype(np.float64), axis=1)
            if not np.isfinite(matrix).all() or not np.isfinite(norms).all() or (norms <= 0).any():
                raise ValueError()
            return matrix
        except (TypeError, ValueError, OverflowError) as error:
            raise ValueError('TM 向量数量、维度或数值异常，未使用零向量代替') from error

    def _embed_batch(self, texts):
        result = self._post(self.embedding, {'model': self.embedding['model_name'], 'input': texts},
                            self.embed_timeout, '向量模型')
        try:
            data = result['data']
            if not isinstance(data, list) or len(data) != len(texts):
                raise ValueError()
            ordered = {}
            for item in data:
                index = item['index']
                if type(index) is not int or not 0 <= index < len(texts) or index in ordered:
                    raise ValueError()
                ordered[index] = item['embedding']
            return self._vectors([ordered[i] for i in range(len(texts))], len(texts))
        except (KeyError, TypeError, ValueError) as error:
            raise ValueError('TM 向量响应不完整或格式异常，请核对模型接口') from error

    def embed(self, texts, progress=None):
        batches = [texts[i:i + self.batch_size] for i in range(0, len(texts), self.batch_size)]
        matrices = [None] * len(batches)
        with ThreadPoolExecutor(max_workers=self.embed_workers) as pool:
            futures = {pool.submit(self._embed_batch, batch): i for i, batch in enumerate(batches)}
            try:
                for done, future in enumerate(as_completed(futures), 1):
                    matrices[futures[future]] = future.result()
                    if progress:
                        progress(done, len(batches))
            except Exception:
                for future in futures:
                    future.cancel()
                raise
        if len({m.shape[1] for m in matrices}) != 1:
            raise ValueError('TM 向量批次维度不一致，请检查模型')
        return self._vectors(np.concatenate(matrices), len(texts))

    def prepare(self, sources, tm_hash, progress=None):
        identity = {'policy': SEMANTIC_POLICY, 'language': self.language.key, 'tm_sha256': tm_hash,
                    'sources_sha256': hashlib.sha256(json.dumps(sources, ensure_ascii=False).encode()).hexdigest(),
                    'model': self.embedding['model_name'],
                    'endpoint_sha256': hashlib.sha256(self.embedding['api_url'].encode()).hexdigest()}
        manifest = json.dumps(identity, sort_keys=True)
        fingerprint = hashlib.sha256(manifest.encode()).hexdigest()
        if self.vectors is not None and getattr(self, '_fingerprint', None) == fingerprint:
            return
        path = self.cache_dir / f'{fingerprint}.npz'
        matrix = None
        if path.exists():
            try:
                with np.load(path, allow_pickle=False) as saved:
                    if saved['manifest'].item() == manifest:
                        matrix = self._vectors(saved['vectors'], len(sources))
            except (OSError, ValueError, KeyError, EOFError, BadZipFile):
                matrix = None
        self.cache_hit = matrix is not None
        if matrix is None:
            matrix = self.embed(sources, progress)
            self.cache_dir.mkdir(parents=True, exist_ok=True)
            temp = path.with_name(f'.{fingerprint}-{uuid.uuid4().hex}.tmp')
            try:
                with temp.open('wb') as stream:
                    np.savez_compressed(stream, vectors=matrix, manifest=np.array(manifest))
                os.replace(temp, path)
            finally:
                temp.unlink(missing_ok=True)
        norms = np.linalg.norm(matrix.astype(np.float64), axis=1, keepdims=True)
        self.vectors = (matrix / norms).astype(np.float32)
        self._fingerprint = fingerprint
        self.sources = sources
        tokenized = [word_tokens(s, self.language.chinese) for s in sources]
        self.bm25 = BM25Okapi(tokenized) if any(tokenized) else None

    def rerank(self, query, documents):
        result = self._post(self.rerank_config, {
            'model': self.rerank_config['model_name'], 'query': query, 'documents': documents,
            'top_n': len(documents), 'input': {'query': query, 'documents': documents},
            'parameters': {'return_documents': False, 'top_n': len(documents)},
        }, self.rank_timeout, '重排模型')
        try:
            items = result.get('results', result.get('output', {}).get('results'))
            if not isinstance(items, list) or len(items) != len(documents):
                raise ValueError()
            scores = {}
            for item in items:
                index, score = item['index'], item['relevance_score']
                if (type(index) is not int or not 0 <= index < len(documents) or index in scores
                        or type(score) not in (int, float) or not math.isfinite(score)):
                    raise ValueError()
                scores[index] = score
            return [scores[i] for i in range(len(documents))]
        except (KeyError, TypeError, ValueError, AttributeError) as error:
            raise ValueError('TM 重排响应不完整或格式异常，未当作“未命中”处理') from error

    def _search(self, query, vector):
        if vector.shape[0] != self.vectors.shape[1]:
            raise ValueError('查询与记忆库向量维度不一致，请更换模型后重新生成缓存')
        semantic = self.vectors @ (vector / np.linalg.norm(vector.astype(np.float64)))
        lexical = self.bm25.get_scores(word_tokens(query, self.language.chinese)) if self.bm25 else np.zeros(len(self.sources))
        maximum = float(np.max(lexical))
        lexical = lexical / maximum if maximum > (0 if self.language.chinese else 5.0) else np.zeros_like(lexical)
        hybrid = 0.7 * semantic + 0.3 * lexical
        indices = np.argsort(-hybrid, kind='stable')[:self.recall_k].tolist()
        scores = self.rerank(query, [self.sources[i] for i in indices])
        ranked = sorted(zip(indices, scores), key=lambda pair: (-pair[1], pair[0]))
        return [(i, float(score), '语义相似') for i, score in ranked if score >= self.threshold][:self.final_k]

    def search_many(self, queries, sources, tm_hash, progress=None):
        report = progress or (lambda percent, message: None)
        self.prepare(sources, tm_hash, lambda done, total: report(
            5 + int(45 * done / total), f'正在生成 TM 向量 [{done}/{total}]（首次运行后缓存）...'))
        report(50, 'TM 向量缓存已载入，正在生成待译原文向量...' if self.cache_hit
               else 'TM 向量已生成，正在生成待译原文向量...')
        vectors = self.embed(queries, lambda done, total: report(
            50 + int(10 * done / total), f'正在生成待译原文向量 [{done}/{total}]...'))
        results = {}
        with ThreadPoolExecutor(max_workers=self.rank_workers) as pool:
            futures = {pool.submit(self._search, query, vector): query for query, vector in zip(queries, vectors)}
            try:
                for done, future in enumerate(as_completed(futures), 1):
                    results[futures[future]] = future.result()
                    report(60 + int(35 * done / len(queries)), f'正在进行 TM 语义召回与重排 [{done}/{len(queries)}]...')
            except Exception:
                for future in futures:
                    future.cancel()
                raise
        return results

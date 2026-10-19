import pandas as pd
import requests
import json
import time
import concurrent.futures


class TermFilter:
    def __init__(self, config_path="config/config.json"):
        # 改为读取 JSON 配置文件
        with open(config_path, 'r', encoding='utf-8') as f:
            cfg = json.load(f).get("llm", {})

        self.api_key = cfg.get("api_key", "")
        self.api_url = cfg.get("api_url", "https://openrouter.ai/api/v1/chat/completions")
        self.model_name = cfg.get("model_name", "deepseek/deepseek-v4.1-flash")
        self.temperature = cfg.get("temperature", 0.1)
        self.max_workers = cfg.get("max_workers", 5)

    def filter_terms_with_llm(self, text, candidates):
        if not candidates:
            return []

        system_prompt = """你是一个专业的 MMORPG 游戏本地化专家，负责筛选机器匹配出的中英对照术语。
任务：验证匹配到的术语是否真的在给定的中文句子中出现，并且语意相符。
规则：
1. 你的回答必须是一个合法的 JSON 对象。
2. JSON 必须包含一个键 "valid_terms"，其值为一个数组，包含所有通过验证的术语字符串。
3. 如果所有术语都不符合，"valid_terms" 应该是一个空数组 []。
4. 严格按照原样输出符合要求的候选词，不要做任何截断或修改。

示例 1：
句子: 受到攻击时有概率使用[蹲下]
候选术语: ["蹲下 (Stoop)", "攻击 (Attack)"]
输出格式: {"valid_terms": ["蹲下 (Stoop)"]}

示例 2：
句子: 力量增加
候选术语: ["力量 (STR)"]
输出格式: {"valid_terms": []}

示例 3：
句子: 艾尔帕兰的冒险者
候选术语: ["艾尔帕兰 (Al De Baran)", "冒险者 (Adventurers)", "旅人 (Traveler)", "W (W)"]
输出格式: {"valid_terms": ["艾尔帕兰 (Al De Baran)", "冒险者 (Adventurers)", "旅人 (Traveler)"]}

请只输出 JSON 对象，不要包含任何其他解释性文本：
{
    "valid_terms": ["词1", "词2"]
}
"""
        user_prompt = f"""句子: {text}\n候选术语: {json.dumps(candidates, ensure_ascii=False)}\n输出格式 JSON:"""

        headers = {
            'Authorization': f'Bearer {self.api_key}',
            'Content-Type': 'application/json'
        }

        payload = {
            "model": self.model_name,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            "response_format": {"type": "json_object"},
            "temperature": self.temperature
        }

        try:
            response = requests.post(
                self.api_url,
                headers=headers,
                json=payload,
                timeout=30,
                proxies={"http": None, "https": None}
            )
            response.raise_for_status()
            result_data = response.json()
            content_str = result_data['choices'][0]['message']['content']

            if content_str.startswith('```json'):
                content_str = content_str.replace('```json', '').replace('```', '').strip()

            parsed_result = json.loads(content_str)
            return parsed_result.get('valid_terms', [])

        except Exception as e:
            print(f"  [大模型接口错误] API请求异常: {e}")
            return candidates

    def _process_single_row(self, index, cn_text, candidates):
        """独立并发方法：处理单行的过滤任务"""
        valid_terms = self.filter_terms_with_llm(cn_text, candidates)
        safe_valid_terms = [t for t in valid_terms if t in candidates]
        invalid_terms = [t for t in candidates if t not in safe_valid_terms]
        return index, ' | '.join(safe_valid_terms), ' | '.join(invalid_terms)

    def process(self, input_file, output_file, task_dict, task_id):
        try:
            task_dict[task_id]['status'] = '正在加载待处理数据...'
            df = pd.read_csv(input_file)

            term_col = '关联术语'
            if '匹配术语' in df.columns:
                term_col = '匹配术语'

            if 'cn' not in df.columns or term_col not in df.columns:
                raise ValueError(f"CSV文件必须包含 'cn' 和 '{term_col}' 列")

            df['生效术语'] = ''
            df['无效术语'] = ''

            valid_tasks = {}
            for index, row in df.iterrows():
                cn_text = str(row['cn'])
                terms_str = str(row[term_col])

                if pd.isna(row[term_col]) or terms_str.strip() == '' or terms_str.lower() == 'nan':
                    continue

                candidates = [t.strip() for t in terms_str.split('|') if t.strip()]
                if candidates:
                    valid_tasks[index] = (cn_text, candidates)

            total = len(df)
            completed = 0
            total_tasks = len(valid_tasks)

            # 使用线程池并发执行大模型过滤
            with concurrent.futures.ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                future_to_idx = {
                    executor.submit(self._process_single_row, idx, cn_text, candidates): idx
                    for idx, (cn_text, candidates) in valid_tasks.items()
                }

                for future in concurrent.futures.as_completed(future_to_idx):
                    idx, active_str, inactive_str = future.result()

                    df.at[idx, '生效术语'] = active_str
                    df.at[idx, '无效术语'] = inactive_str

                    completed += 1
                    task_dict[task_id]['progress'] = int((completed / max(total_tasks, 1)) * 100)
                    task_dict[task_id]['status'] = f'正在并发过滤术语 [{completed}/{total_tasks}]...'

            if term_col in df.columns:
                df = df.drop(columns=[term_col])

            df.to_csv(output_file, index=False, encoding='utf-8-sig')

            task_dict[task_id]['progress'] = 100
            task_dict[task_id]['status'] = f'分析完成！共处理 {total} 条数据'
            task_dict[task_id]['state'] = 'completed'
            task_dict[task_id]['result_file'] = output_file

        except Exception as e:
            task_dict[task_id]['status'] = f'处理失败: {str(e)}'
            task_dict[task_id]['state'] = 'error'
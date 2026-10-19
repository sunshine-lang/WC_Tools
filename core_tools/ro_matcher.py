# core_tools/ro_matcher.py
import pandas as pd
import requests
from bs4 import BeautifulSoup
from urllib.parse import quote
import time
import re


class RODataMatcher:
    def __init__(self):
        self.en_dict = {
            'item': None, 'card': None, 'monster': None,
            'npc': None, 'map': None, 'skill': None
        }

    def get_id_from_dvg(self, name, search_type):
        encoded_name = quote(name, encoding='utf-8')
        endpoints = {
            'item': 'items.php',
            'card': 'cards.php',
            'monster': 'monster.php',
            'npc': 'npcs.php',
            'map': 'map.php',
            'skill': 'skill.php'
        }
        endpoint = endpoints.get(search_type, 'items.php')
        url = f"https://ro.dvg.cn/{endpoint}?words={encoded_name}"
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        try:
            response = requests.get(url, headers=headers, timeout=10)
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.text, 'html.parser')
            candidates = []
            for link_tag in soup.find_all('a'):
                link_text = link_tag.get_text(strip=True)
                if not link_text:
                    continue
                clean_link_text = re.sub(r'\s*\[.*?\]$', '', link_text)
                score = 0
                if link_text == name or clean_link_text == name:
                    score = 100
                else:
                    if clean_link_text.endswith(name) or clean_link_text.endswith(name + " "):
                        score = 80
                    elif name in clean_link_text:
                        score = 50
                if score > 0:
                    entity_id = None
                    for attr in ['data-itemid', 'data-skillid', 'data-mobid', 'data-npcid', 'data-mapid']:
                        if link_tag.has_attr(attr):
                            entity_id = link_tag.get(attr)
                            break
                    if not entity_id:
                        href = link_tag.get('href', '')
                        match = re.search(r'(?:id|map|name)=([^&]+)', href)
                        if match:
                            entity_id = match.group(1)
                    if entity_id:
                        if search_type == 'map':
                            if '_' not in entity_id and not any(c.isdigit() for c in entity_id):
                                score += 50
                            if any(kw in clean_link_text for kw in [' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ']):
                                score += 5
                        candidates.append({
                            'id': entity_id,
                            'score': score,
                            'length': len(clean_link_text)
                        })
            if candidates:
                candidates.sort(key=lambda x: (-x['score'], x['length']))
                return candidates[0]['id']
            return None
        except Exception as e:
            print(f"      [DVG查询失败]: {e}")
            return None

    def load_official_db(self, search_type):
        print(f"\n[预加载] 正在从 GitHub 抓取 {search_type} 官服英文数据...")
        db_mapping = {
            'skill': [
                "https://raw.githubusercontent.com/rathena/rathena/master/db/re/skill_db.yml"
            ],
            'item': [
                "https://raw.githubusercontent.com/rathena/rathena/master/db/re/item_db_equip.yml",
                "https://raw.githubusercontent.com/rathena/rathena/master/db/re/item_db_usable.yml",
                "https://raw.githubusercontent.com/rathena/rathena/master/db/re/item_db_etc.yml"
            ],
            'card': [
                "https://raw.githubusercontent.com/rathena/rathena/master/db/re/item_db_equip.yml",
                "https://raw.githubusercontent.com/rathena/rathena/master/db/re/item_db_usable.yml",
                "https://raw.githubusercontent.com/rathena/rathena/master/db/re/item_db_etc.yml"
            ],
            'monster': [
                "https://raw.githubusercontent.com/rathena/rathena/master/db/re/mob_db.yml"
            ]
        }
        if search_type not in db_mapping:
            return {}
        urls = db_mapping[search_type]
        if search_type == 'skill':
            name_regex = r'(?:^|\n)\s*Description:\s*(.+)'
        else:
            name_regex = r'(?:^|\n)\s*Name:\s*(.+)'
        data_dict = {}
        for url in urls:
            try:
                response = requests.get(url, timeout=20)
                if response.status_code != 200:
                    continue
                for block in response.text.split("- Id: ")[1:]:
                    id_match = re.match(r'^(\d+)', block)
                    desc_match = re.search(name_regex, block)
                    if id_match and desc_match:
                        data_dict[id_match.group(1)] = desc_match.group(1).strip().strip('"').strip("'")
            except Exception as e:
                print(f"      [GitHub抓取失败]: {e}")
                continue
        return data_dict

    def process(self, input_file, output_file, search_type, task_dict, task_id):
        try:
            if self.en_dict[search_type] is None:
                task_dict[task_id]['status'] = '正在加载GitHub官服数据...'
                self.en_dict[search_type] = self.load_official_db(search_type)
            task_dict[task_id]['status'] = '正在读取CSV文件...'
            df = pd.read_csv(input_file)
            if 'cn' not in df.columns:
                raise ValueError("CSV文件中必须包含'cn'列")
            if 'id' not in df.columns: df['id'] = pd.Series(dtype='object')
            if 'en' not in df.columns: df['en'] = pd.Series(dtype='object')

            # 【新增】初始化 URL 列 和 URL_RE 列
            if 'URL' not in df.columns: df['URL'] = pd.Series(dtype='object')
            if 'URL_RE' not in df.columns: df['URL_RE'] = pd.Series(dtype='object')

            total = len(df)
            for idx, row in df.iterrows():
                cn_name = str(row['cn']).strip()
                if not cn_name or cn_name == 'nan': continue
                if search_type == 'card' and not cn_name.endswith('卡片'):
                    cn_name += '卡片'
                    df.at[idx, 'cn'] = cn_name
                task_dict[task_id]['progress'] = int((idx / total) * 100)
                task_dict[task_id]['status'] = f'[{idx + 1}/{total}] 正在查询: {cn_name}'
                entity_id = self.get_id_from_dvg(cn_name, search_type)
                time.sleep(1)

                if not entity_id:
                    df.at[idx, 'id'], df.at[idx, 'en'] = "无", "无"
                    # 如果没有查到数据，URL 给空字符串
                    df.at[idx, 'URL'] = ""
                    df.at[idx, 'URL_RE'] = ""
                    continue

                if search_type in ['map', 'npc']:
                    en_name = entity_id
                else:
                    en_name = self.en_dict[search_type].get(entity_id, "暂无英文名")

                df.at[idx, 'id'] = entity_id
                df.at[idx, 'en'] = en_name

                # 生成正常 URL 和 带有 re_ 的 Renewal URL
                url = ""
                url_re = ""
                if search_type == 'item':
                    url = f"https://ratemyserver.net/index.php?iname={en_name}&page=item_db&quick=1&isearch=Search"
                    url_re = f"https://ratemyserver.net/index.php?iname={en_name}&page=re_item_db&quick=1&isearch=Search"
                elif search_type == 'monster':
                    url = f"https://ratemyserver.net/index.php?mob_name={en_name}&page=mob_db&quick=1&f=1&mob_search=Search"
                    url_re = f"https://ratemyserver.net/index.php?mob_name={en_name}&page=re_mob_db&quick=1&f=1&mob_search=Search"
                elif search_type == 'card':
                    url = f"https://ratemyserver.net/index.php?iname={en_name}&page=item_db&quick=1&isearch=Search"
                    url_re = f"https://ratemyserver.net/index.php?iname={en_name}&page=re_item_db&quick=1&isearch=Search"
                elif search_type == 'skill':
                    url = f"https://ratemyserver.net/index.php?page=skill_db&skid={entity_id}"
                    url_re = f"https://ratemyserver.net/index.php?page=re_skill_db&skid={entity_id}"

                df.at[idx, 'URL'] = url
                df.at[idx, 'URL_RE'] = url_re

            df.to_csv(output_file, index=False, encoding='utf-8-sig')
            task_dict[task_id]['progress'] = 100
            task_dict[task_id]['status'] = f'处理完成，共 {total} 条数据'
            task_dict[task_id]['state'] = 'completed'
            task_dict[task_id]['result_file'] = output_file
        except Exception as e:
            task_dict[task_id]['status'] = f'发生错误: {str(e)}'
            task_dict[task_id]['state'] = 'error'
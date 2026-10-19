import pandas as pd
import re


class TagCleaner:
    def process(self, input_file, output_file, task_dict, task_id):
        try:
            task_dict[task_id]['status'] = '读取文件中...'
            df = pd.read_csv(input_file)

            if 'cn' not in df.columns:
                raise ValueError("上传的 CSV 必须包含 'cn' 列")

            total = len(df)
            for idx, row in df.iterrows():
                if idx % 50 == 0 or idx == total - 1:
                    task_dict[task_id]['progress'] = int((idx / total) * 100)
                    task_dict[task_id]['status'] = f'正在清洗标签空格 [{idx + 1}/{total}]...'

                cn_text = str(row.get('cn', ''))
                if cn_text and cn_text != 'nan':
                    # 💡 终极优化版：移除任何标签（无论开闭）左右两侧的所有多余空格，且绝不触碰尖括号内部
                    cleaned_text = re.sub(r'\s*(<[^>]+>)\s*', r'\1', cn_text)

                    df.at[idx, 'cn'] = cleaned_text

            df.to_csv(output_file, index=False, encoding='utf-8-sig')
            task_dict[task_id]['progress'] = 100
            task_dict[task_id]['status'] = f'处理完成，共清洗 {total} 条数据'
            task_dict[task_id]['state'] = 'completed'
            task_dict[task_id]['result_file'] = output_file

        except Exception as e:
            task_dict[task_id]['status'] = f'处理错误: {str(e)}'
            task_dict[task_id]['state'] = 'error'
import pandas as pd
import json
import io
import os
import re


def generate_organized_excels_smart(file_path, folder_name='output_excels'):
    # 1. 출력 폴더 생성
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
        print(f"📂 폴더 확인: {folder_name}")

    # 2. 텍스트 파일 읽기
    if not os.path.exists(file_path):
        print(f"❌ 파일을 찾을 수 없습니다: {file_path}")
        return

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            raw_content = f.read().strip()

        # JSON 파싱 전처리
        clean_content = raw_content.replace('```json', '').replace('```', '').strip()

        # 콤마 보정 등 JSON 파싱 시도
        try:
            data = json.loads(clean_content)
        except json.JSONDecodeError:
            # 콤마 누락 자동 보정
            fixed_content = re.sub(r'(?<=[^\\]")\s+(?=")', ',\n', clean_content)
            try:
                data = json.loads(fixed_content)
            except:
                # 최악의 경우: JSON 형식이 아니더라도 텍스트 덩어리로 처리 시도
                data = {"merged_data": clean_content}

    except Exception as e:
        print(f"❌ 파일 읽기 에러: {e}")
        return

    # 3. 모든 데이터 하나로 합치기 (Master DataFrame 생성)
    all_dfs = []

    for key, text_data in data.items():
        try:
            # 헤더 찾기 로직
            lines = str(text_data).strip().split('\n')
            start_index = -1
            for i, line in enumerate(lines):
                if 'pdffile' in line and '\t' in line:
                    start_index = i
                    break

            if start_index != -1:
                clean_text = '\n'.join(lines[start_index:])
                df_chunk = pd.read_csv(io.StringIO(clean_text), sep='\t')
                all_dfs.append(df_chunk)
        except Exception as e:
            print(f"⚠️ 데이터 병합 중 경고 ({key}): {e}")

    if not all_dfs:
        print("❌ 처리할 데이터가 없습니다.")
        return

    # 전체 데이터를 하나의 프레임으로 병합
    master_df = pd.concat(all_dfs, ignore_index=True)

    # 4. 'pdffile' 컬럼 기준으로 다시 쪼개서 엑셀 저장 (핵심 로직)
    if 'pdffile' not in master_df.columns:
        print("❌ 데이터에 'pdffile' 컬럼이 없어 분리할 수 없습니다.")
        return

    unique_files = master_df['pdffile'].unique()
    print(f"🔍 총 {len(unique_files)}개의 고유 파일을 발견했습니다. 분리 저장을 시작합니다...")

    success_count = 0
    for pdf_filename in unique_files:
        try:
            # 해당 파일의 데이터만 필터링
            file_df = master_df[master_df['pdffile'] == pdf_filename]

            # 파일명 정제 (.pdf 제거 등)
            base_name = str(pdf_filename).replace('.pdf', '').replace('.png', '').strip()
            # 파일명에 엑셀에서 못 쓰는 특수문자가 있다면 제거/변경
            base_name = re.sub(r'[\\/*?:"<>|]', "_", base_name)

            excel_filename = base_name + ".xlsx"
            excel_path = os.path.join(folder_name, excel_filename)

            # 엑셀 저장
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                categories = {
                    'Income Statement': 'Income Statement',
                    'Balance Sheet': 'Balance Sheet',
                    'Working Capital': 'Working Capital'
                }

                found_any_sheet = False
                for cat_key, sheet_name in categories.items():
                    # 해당 카테고리 데이터 필터링
                    sheet_df = file_df[file_df['table'].str.contains(cat_key, case=False, na=False)]

                    if not sheet_df.empty:
                        sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        found_any_sheet = True

                if found_any_sheet:
                    print(f"✅ 생성 완료: {excel_filename}")
                    success_count += 1
                else:
                    print(f"⚠️ 데이터 부족으로 생성 건너뜀: {excel_filename}")

        except Exception as e:
            print(f"❌ '{pdf_filename}' 저장 실패: {e}")

    print(f"\n🎉 총 {success_count}개의 엑셀 파일이 완벽하게 분리 생성되었습니다!")


if __name__ == "__main__":
    generate_organized_excels_smart('output_gemini.txt')
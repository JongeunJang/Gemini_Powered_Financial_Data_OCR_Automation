import pandas as pd
import json
import io
import os
import re


def generate_organized_excels_final(file_path, folder_name='output_excels'):
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

        # JSON 파싱 전처리 (마크다운 제거)
        clean_content = raw_content.replace('```json', '').replace('```', '').strip()

        # 콤마 누락 자동 보정 로직
        try:
            data = json.loads(clean_content)
        except json.JSONDecodeError:
            print("⚠️ JSON 문법 보정 중...")
            fixed_content = re.sub(r'(?<=[^\\]")\s+(?=")', ',\n', clean_content)
            try:
                data = json.loads(fixed_content)
            except:
                print("❌ JSON 파싱 실패. 텍스트 파일의 형식을 확인해주세요.")
                return

    except Exception as e:
        print(f"❌ 파일 읽기 에러: {e}")
        return

    # 3. 데이터 처리 및 엑셀 생성
    success_count = 0
    for pdf_name, v19_text in data.items():
        try:
            # === 진짜 데이터 시작점(Header) 찾기 ===
            lines = v19_text.strip().split('\n')
            start_index = -1

            for i, line in enumerate(lines):
                if 'pdffile' in line and '\t' in line:
                    start_index = i
                    break

            if start_index == -1:
                print(f"⚠️ '{pdf_name}' 건너뜀: 유효한 데이터 헤더(pdffile)를 찾을 수 없습니다.")
                continue

            # 진짜 데이터만 다시 합치기
            clean_table_text = '\n'.join(lines[start_index:])

            # 데이터프레임 변환
            df = pd.read_csv(io.StringIO(clean_table_text), sep='\t')

            # .pdf 확장자 보정
            if not pdf_name.lower().endswith('.pdf'):
                pdf_name += ".pdf"

            # 엑셀 저장 경로 설정
            excel_filename = pdf_name.replace('.pdf', '') + ".xlsx"
            excel_path = os.path.join(folder_name, excel_filename)

            # 엑셀 저장 (시트 분리 로직)
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                categories = {
                    'Income Statement': 'Income Statement',
                    'Balance Sheet': 'Balance Sheet',
                    'Working Capital': 'Working Capital'
                }

                found_any = False
                for cat_key, sheet_name in categories.items():
                    # 대소문자 구분 없이 포함 여부 확인
                    filtered_df = df[df['table'].str.contains(cat_key, case=False, na=False)]
                    if not filtered_df.empty:
                        filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        found_any = True

                # [삭제됨] Etc 시트 생성 로직을 제거했습니다.

            if found_any:
                print(f"✅ 생성 완료: {excel_filename}")
                success_count += 1
            else:
                # 3가지 카테고리에 해당하는 데이터가 하나도 없으면 파일은 생성되지만 빈 껍데기일 수 있음
                # 혹은 openpyxl 특성상 기본 시트가 하나 생길 수 있음
                print(f"⚠️ 경고: {excel_filename} (I/S, B/S, W/C 데이터가 발견되지 않음)")
                # 필요시 여기서 os.remove(excel_path)를 호출하여 빈 파일 삭제 가능

        except Exception as e:
            print(f"❌ '{pdf_name}' 처리 중 오류: {e}")

    print(f"\n🎉 총 {success_count}개의 엑셀 파일이 생성되었습니다!")


if __name__ == "__main__":
    generate_organized_excels_final('gemini_output.txt')
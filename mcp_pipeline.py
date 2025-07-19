# mcp_pipeline_openrouter.py
import os
import zipfile
import xml.etree.ElementTree as ET
import win32com.client as win32
import openai
import json
from dotenv import load_dotenv

# --- .env 파일에서 환경 변수 로드 ---
load_dotenv()

# --- 1. HWP 변환기 (HWP -> HWPX) ---
# (이전 코드와 동일, 변경 없음)
def convert_hwp_to_hwpx(hwp_path, output_folder):
    """HWP 파일을 HWPX 파일로 변환합니다."""
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        filename = os.path.basename(hwp_path)
        hwpx_filename = os.path.splitext(filename)[0] + ".hwpx"
        hwpx_path = os.path.join(output_folder, hwpx_filename)
        hwp.Open(hwp_path)
        hwp.SaveAs(hwpx_path, "HWPX")
        hwp.Quit()
        print(f"  - 변환 성공: {filename} -> {hwpx_filename}")
        return hwpx_path
    except Exception as e:
        print(f"  - 변환 오류: {hwp_path} 처리 중 문제 발생 - {e}")
        return None

# --- 2. HWPX 해석기 (HWPX -> Markdown) - [수정됨] ---
def hwpxto_markdown_parser(hwpx_path):
    """HWPX 파일의 텍스트와 표를 마크다운 형식으로 추출합니다."""
    # (표 변환 로직은 이전과 동일)
    def convert_table_to_markdown(tbl_element, ns):
        rows_data = []
        for tr in tbl_element.findall('hwp:tr', ns):
            cell_data = [''.join(tc.itertext()).strip().replace('\n', ' ') for tc in tr.findall('hwp:tc', ns)]
            rows_data.append(cell_data)
        if not rows_data or not rows_data[0]: return ""
        num_columns = len(rows_data[0])
        header = "| " + " | ".join(rows_data[0]) + " |"
        separator = "| " + " | ".join(["---"] * num_columns) + " |"
        body_rows = "\n".join(["| " + " | ".join(row) + " |" for row in rows_data[1:]])
        return f"\n{header}\n{separator}\n{body_rows}\n"

    try:
        zip_ref = zipfile.ZipFile(hwpx_path, 'r')
        content_xml_path = 'Contents/section0.xml'
        xml_content = zip_ref.read(content_xml_path)
        root = ET.fromstring(xml_content)
        ns = {'hwp': 'http://www.hancom.co.kr/hwpml/2010/namespace'}
        
        # --- ★★★ 여기가 수정된 부분 ★★★ ---
        body = root.find('.//hwp:body', ns)
        
        # body가 None인지 확인하는 방어 코드 추가
        if body is None:
            print(f"  - 해석 오류: {os.path.basename(hwpx_path)} 파일에서 'body' 섹션을 찾을 수 없습니다. 건너뜁니다.")
            return None
        # --- ★★★ 수정 끝 ★★★ ---
            
        content_list = []
        for element in body: # 이제 body가 None이 아님이 보장됨
            if element.tag.endswith("}p"):
                content_list.append(''.join(element.itertext()).strip())
            elif element.tag.endswith("}tbl"):
                content_list.append(convert_table_to_markdown(element, ns))
        
        print(f"  - 해석 성공: {os.path.basename(hwpx_path)}")
        return "\n\n".join(content_list)
    except Exception as e:
        print(f"  - 해석 오류: {hwpx_path} 처리 중 예기치 않은 문제 발생 - {e}")
        return None

# --- 3. LLM 지능부 (OpenRouter 연동) ---
def get_modification_plan_from_llm(document_content, user_request):
    """OpenRouter의 Mistral 모델을 호출하여 문서 수정 계획(JSON)을 받습니다."""
    print("  - LLM 호출 (OpenRouter): 수정 계획을 요청합니다...")
    
    # .env 파일에서 API 키를 불러옵니다.
    api_key = os.getenv("OPENROUTER_API_KEY")
    if not api_key:
        print("  - LLM 오류: .env 파일에 OPENROUTER_API_KEY가 설정되지 않았습니다.")
        return None

    # OpenAI 라이브러리를 OpenRouter API와 호환되도록 클라이언트 설정
    client = openai.OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key,
    )
    
    # Mistral 모델에 최적화된 시스템 프롬프트
    system_prompt = "You are an extremely competent document editing expert. Respond only in JSON format according to the user's request. Your task is to create a list of 'replace_text' actions."
    
    user_prompt = f"""Based on the [USER_REQUEST] below, create a specific 'modification plan' for the [ORIGINAL_DOCUMENT].
The plan must be a JSON object containing a single key "actions", which is a list of replacement operations.
Each operation should be an object with two keys: "find" (the exact text to find in the original document) and "replace" (the text to replace it with).

[USER_REQUEST]
{user_request}

[ORIGINAL_DOCUMENT]
{document_content}
"""
    
    try:
        response = client.chat.completions.create(
            model="mistralai/mistral-7b-instruct:free", # 무료 모델 사용, mistral-small-3.2-24b-instruct:free 사용도 가능
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            response_format={"type": "json_object"},
            temperature=0.1,  # Mistral 모델은 낮은 온도를 권장합니다 [4]
        )
        plan = json.loads(response.choices[0].message.content)
        print("  - LLM 응답: 수정 계획 수신 완료.")
        return plan.get('actions', [])
    except Exception as e:
        print(f"  - LLM 오류: API 호출 중 문제 발생 - {e}")
        return None

# --- 4. HWP 제어기 (수정 계획 실행 및 파일 생성) ---
# (이전 코드와 동일, 변경 없음)
def execute_modifications_and_save(original_hwpx_path, actions, output_folder):
    """수정 계획에 따라 문서를 수정하고 최종 HWP 파일로 저장합니다."""
    print(f"  - 최종 문서 생성 시작: {os.path.basename(original_hwpx_path)}")
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.Open(original_hwpx_path)
        for action in actions:
            if action['action'] == 'replace_text':
                hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
                option = hwp.HParameterSet.HFindReplace
                option.FindString = action['find']
                option.ReplaceString = action['replace']
                option.IgnoreMessage = 1
                hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
        filename = os.path.basename(original_hwpx_path)
        output_filename = os.path.splitext(filename)[0] + "_수정본.hwp"
        save_path = os.path.join(output_folder, output_filename)
        hwp.SaveAs(save_path)
        hwp.Quit()
        print(f"  - 생성 완료: {output_filename}")
        return save_path
    except Exception as e:
        print(f"  - 생성 오류: 최종 파일 저장 중 문제 발생 - {e}")
        return None

# --- 5. MCP 메인 파이프라인 ---
# (이전 코드와 동일, 변경 없음)
def run_mcp_pipeline(folder_path, user_request):
    """지정된 폴더와 사용자 요청에 따라 전체 자동화 파이프라인을 실행합니다."""
    print("="*20 + " MCP 파이프라인 시작 " + "="*20)
    hwpx_temp_folder = os.path.join(folder_path, "temp_hwpx")
    output_folder = os.path.join(folder_path, "final_output")
    if not os.path.exists(hwpx_temp_folder): os.makedirs(hwpx_temp_folder)
    if not os.path.exists(output_folder): os.makedirs(output_folder)
    all_files = os.listdir(folder_path)
    for filename in all_files:
        file_path = os.path.join(folder_path, filename)
        if os.path.isdir(file_path): continue
        print(f"\n[파일 처리 시작] '{filename}'")
        file_ext = os.path.splitext(filename)[1].lower()
        if file_ext == '.hwp':
            hwpx_path = convert_hwp_to_hwpx(file_path, hwpx_temp_folder)
        elif file_ext == '.hwpx':
            hwpx_path = file_path
        else:
            print("  - 지원하지 않는 파일 형식입니다. 건너뜁니다.")
            continue
        if not hwpx_path or not os.path.exists(hwpx_path): continue
        markdown_content = hwpxto_markdown_parser(hwpx_path)
        if not markdown_content: continue
        modification_plan = get_modification_plan_from_llm(markdown_content, user_request)
        if not modification_plan: continue
        execute_modifications_and_save(hwpx_path, modification_plan, output_folder)
    print("\n" + "="*20 + " MCP 파이프라인 종료 " + "="*20)


# --- 스크립트 실행 부분 ---
if __name__ == "__main__":
    target_folder = "C:\\python\\HWP\\targets"
    request = "문서에 있는 표의 초4학년의 일자를 다음 주 목요일 1~2교시로 바꾸고 교과별 평가 범위에 인사를 간단하게 써 줘. 문항 수는 13으로 채워주고."
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
        print(f"'{target_folder}' 폴더를 생성했습니다. 이 폴더에 HWP/HWPX 파일을 넣고 스크립트를 다시 실행해주세요.")
    else:
        run_mcp_pipeline(target_folder, request)

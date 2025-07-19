import sys
import os
import subprocess
import shutil
import time
import psutil
from dotenv import load_dotenv

# 프로젝트 루트 경로 설정
project_root = os.path.abspath(os.path.dirname(__file__))
sys.path.append(os.path.join(project_root, 'hwp-mcp', 'src', 'tools'))

from hwp_controller import HwpController  # import 확인

# .env 파일 로드
load_dotenv()
os.environ["GEMINI_API_KEY"] = os.getenv("gemini_API_KEY")

def call_gemini_cli(prompt, input_content=""):
    command = ["gemini", "-m", "gemini-2.5-pro", "-p", prompt]
    try:
        process = subprocess.run(
            command, input=input_content.encode('utf-8'),
            capture_output=True, check=True, timeout=600, shell=True
        )
        return process.stdout.decode('utf-8').strip()
    except Exception as e:
        raise Exception(f"Gemini CLI 실행 오류: {e}")

def organize_hwp_file(file_path, user_prompt="아래는 MCP로 추출된 HWP 실제 내용이야. 이 텍스트를 기반으로 불필요한 부분 제거 후 3줄로 요약해. 추가 입력 요구 금지. 직접 요약만 출력."):
    # 공유 오류 방지: 파일 복사
    temp_path = file_path + ".temp"
    shutil.copyfile(file_path, temp_path)
    abs_file_path = os.path.abspath(temp_path)
    print(f"사용 중 임시 파일 경로: {abs_file_path}")
    
    hwp = HwpController()
    if not hwp.connect(visible=False):
        raise Exception("HWP 연결 실패")

    if not hwp.open_document(abs_file_path):
        raise Exception("open_document 반환 False: 실패 로그 확인")

    # 내용 추출 (MCP 직접 사용) + 로그 출력
    hwp_content = hwp.get_text()
    print(f"디버깅: 추출된 HWP 내용 (길이: {len(hwp_content)}):\n{hwp_content[:200]}...")  # 처음 200자 출력 (전체 확인용)
    if not hwp_content.strip():  # 빈 내용 확인
        raise Exception("텍스트 추출 실패: 내용이 비어 있음. 파일 확인하세요.")

    # 모델 호출: full_prompt에 내용 포함 (input_content 비사용, 안정성 위해)
    full_prompt = f"{user_prompt}\n\n추출 내용:\n{hwp_content}"
    response = call_gemini_cli(full_prompt)

    # 수정 예시: 모델 응답 기반
    hwp.replace_text("불필요", "", replace_all=True)  # 예시 정리

    # 저장 전 권한 설정
    try:
        os.chmod(file_path, 0o666)
        print("디버깅: 원본 파일 권한 설정 성공")
    except Exception as e:
        print(f"권한 설정 실패: {e}")

    # 저장
    try:
        hwp.save_document(file_path)
        print("디버깅: 저장 성공")
    except Exception as e:
        raise Exception(f"문서 저장 실패: {e}")

    # 종료: disconnect 후 HWP 프로세스 강제 종료
    hwp.disconnect()
    print("디버깅: 연결 종료 성공")
    time.sleep(2)
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == 'Hwp.exe':
            proc.terminate()
            print("디버깅: HWP 프로세스 강제 종료")

    # 임시 파일 삭제: 반복 시도
    max_attempts = 5
    for attempt in range(max_attempts):
        try:
            os.remove(temp_path)
            print("디버깅: 임시 파일 삭제 성공")
            break
        except PermissionError as e:
            print(f"삭제 시도 {attempt+1}/{max_attempts} 실패: {e} - 2초 대기 후 재시도")
            time.sleep(2)
    else:
        print("임시 파일 삭제 최종 실패 - 수동 삭제 필요: " + temp_path)

    return "HWP 파일 정리 완료: " + response

# 실행 예시
if __name__ == "__main__":
    file_path = "./targets/sample.hwp"
    result = organize_hwp_file(file_path)
    print(result)

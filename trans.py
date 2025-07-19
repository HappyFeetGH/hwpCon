import os
import win32com.client as win32

def convert_hwp_to_hwpx_in_folder(folder_path):
    """
    지정된 폴더 내의 모든 .hwp 파일을 .hwpx 파일로 자동 변환합니다.
    """
    print(f"'{folder_path}' 폴더에서 HWP 파일 변환을 시작합니다...")
    
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    # 보안 팝업 우회 (이전에 설정한 레지스트리 방식과 함께 사용하면 더 안정적)
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    
    # 변환된 파일들을 저장할 하위 폴더 생성
    output_folder = os.path.join(folder_path, "converted_hwpx")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    try:
        files = os.listdir(folder_path)
        hwp_files = [f for f in files if f.lower().endswith(".hwp")]

        if not hwp_files:
            print("변환할 HWP 파일이 없습니다.")
            return

        for filename in hwp_files:
            hwp_path = os.path.join(folder_path, filename)
            # 출력 파일명 설정 (예: original.hwp -> original.hwpx)
            hwpx_filename = os.path.splitext(filename)[0] + ".hwpx"
            hwpx_path = os.path.join(output_folder, hwpx_filename)
            
            print(f"  - 변환 중: {filename} -> {hwpx_filename}")
            
            # 한글 프로그램으로 파일 열기
            hwp.Open(hwp_path)
            # 다른 이름으로 저장 기능을 이용해 HWPX 포맷으로 저장
            hwp.SaveAs(hwpx_path, "HWPX")

        print(f"\n✅ 변환 완료! 모든 파일이 '{output_folder}' 폴더에 저장되었습니다.")

    except Exception as e:
        print(f"오류 발생: {e}")
    finally:
        hwp.Quit()

# --- 이 코드를 실행하는 예시 ---
if __name__ == "__main__":
    # 상급 기관에서 HWP 파일들을 다운로드 받은 폴더를 지정
    source_folder = "C:/python/HWP/targets" 
    
    # 예시를 위해 임시 폴더와 파일을 만듭니다.
    if not os.path.exists(source_folder):
        os.makedirs(source_folder)
        # 임시 hwp 파일을 하나 만들어 둡니다. (실제로는 이 파일이 이미 있어야 함)
        # 이 부분은 테스트를 위한 것이므로, 실제 hwp 파일을 폴더에 넣어두고 실행하세요.
        # hwp_temp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        # hwp_temp.SaveAs(os.path.join(source_folder, "test.hwp"))
        # hwp_temp.Quit()
        print(f"'{source_folder}' 폴더가 생성되었습니다. 이 폴더에 HWP 파일을 넣고 다시 실행해주세요.")


    convert_hwp_to_hwpx_in_folder(source_folder)


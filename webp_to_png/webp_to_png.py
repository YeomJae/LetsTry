import os
from PIL import Image
from tkinter import Tk
from tkinter.filedialog import askdirectory

def convert_webp_to_png(input_dir, output_dir):
    # 입력 디렉토리 내 모든 파일 탐색
    for filename in os.listdir(input_dir):
        if filename.lower().endswith(".webp"):
            webp_path = os.path.join(input_dir, filename)
            png_filename = filename.rsplit('.', 1)[0] + ".png"
            png_path = os.path.join(output_dir, png_filename)
            
            # webp 파일 열기 및 png로 저장
            with Image.open(webp_path) as img:
                img.save(png_path, "PNG")
            print(f"변환 완료: {webp_path} -> {png_path}")

if __name__ == "__main__":
    # GUI를 통해 폴더 선택
    root = Tk()
    root.withdraw()  # Tkinter 창 숨기기
    
    # 입력 폴더 선택
    input_directory = askdirectory(title="webp 파일이 있는 폴더를 선택하세요")
    if not input_directory:
        print("입력 폴더를 선택하지 않았습니다.")
        exit()
    
    # 저장 폴더 선택
    output_directory = askdirectory(title="변환된 png 파일을 저장할 폴더를 선택하세요")
    if not output_directory:
        print("저장 폴더를 선택하지 않았습니다.")
        exit()
    
    # 변환 시작
    convert_webp_to_png(input_directory, output_directory)
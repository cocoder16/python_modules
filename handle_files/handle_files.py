import os


def remove_file(path_str: str):
    # 이 파일이 있으면 자동로그인 설정이 활성화된다.

    if os.path.isfile(path_str):
        os.remove(path_str)

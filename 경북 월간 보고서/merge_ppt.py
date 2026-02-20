from io import BytesIO
from pptx_merger import Merger

def merge_ppts_with_merger(pptx_paths, output_path):
    """
    pptx_paths: 합칠 pptx 파일 경로 리스트 (예: 3개)
    output_path: 결과 pptx 경로
    """
    merger = Merger()

    # 1) 파일들을 BytesIO로 읽기
    src_docs = []
    for path in pptx_paths:
        with open(path, "rb") as f:
            src_docs.append(BytesIO(f.read()))

    # 2) 병합 (전체 슬라이드)
    merged_io = merger.merge_slides(src_docs)

    # 3) 결과를 파일로 저장
    with open(output_path, "wb") as f:
        f.write(merged_io.getvalue())

    print(f"template 병합 완료: {output_path}")

    

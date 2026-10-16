import datetime
import io
import json
import os
import re
from typing import Dict, List, Optional, Tuple

import streamlit as st
from dotenv import load_dotenv
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# .env 환경 변수 자동 로드
load_dotenv()

# ==========================================
# 1. 페이지 환경설정 및 커스텀 스타일링
# ==========================================
st.set_page_config(
    page_title="신입사원 독서경영 IDEA SHEET 자동 생성기",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-title {
        font-size: 2.1rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 1.05rem;
        color: #4B5563;
        margin-bottom: 1.5rem;
    }
    .card-box {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1.2rem;
        margin-bottom: 1rem;
    }
    .badge-success {
        display: inline-block;
        background-color: #DCFCE7;
        color: #166534;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.88rem;
    }
    .badge-warning {
        display: inline-block;
        background-color: #FEF3C7;
        color: #92400E;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.88rem;
    }
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==========================================
# 2. Gemini API 호출 유틸리티
# ==========================================
def call_gemini_api(api_key: str, prompt: str, model_name: str = "gemini-3.8-flash") -> str:
    """
    google-genai 최신 SDK를 우선 호출하고, 실패 시 google-generativeai로 fallback 처리합니다.
    어떤 예외도 조용히 삼키지 않고 상세 원인을 반환합니다.
    """
    if not api_key:
        raise ValueError("Google Gemini API 키가 입력되지 않았습니다. 사이드바에서 API 키를 입력해 주세요.")

    collected_errors = []

    # 1. 최신 google-genai 라이브러리 시도
    try:
        from google import genai

        client = genai.Client(api_key=api_key)
        try:
            response = client.models.generate_content(model=model_name, contents=prompt)
            if response and response.text:
                return response.text
            else:
                collected_errors.append(f"google-genai ({model_name}): 빈 텍스트 응답이 반환되었습니다.")
        except Exception as e:
            collected_errors.append(f"google-genai ({model_name}) 호출 실패: {type(e).__name__} - {str(e)}")
            # 지정된 모델 호출 실패 시 gemini-3.8-flash로 재시도 (모델명이 다를 경우에만)
            if model_name != "gemini-3.8-flash":
                try:
                    retry_res = client.models.generate_content(model="gemini-3.8-flash", contents=prompt)
                    if retry_res and retry_res.text:
                        return retry_res.text
                except Exception as retry_e:
                    collected_errors.append(f"google-genai (gemini-3.8-flash 재시도) 실패: {type(retry_e).__name__} - {str(retry_e)}")
    except ImportError:
        collected_errors.append("google-genai 패키지가 미설치되어 있습니다.")
    except Exception as e:
        collected_errors.append(f"google-genai 클라이언트 초기화 실패: {type(e).__name__} - {str(e)}")

    # 2. 레거시 google.generativeai 라이브러리 시도
    try:
        import google.generativeai as legacy_genai

        legacy_genai.configure(api_key=api_key)
        try:
            model = legacy_genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            if response and response.text:
                return response.text
            else:
                collected_errors.append(f"google-generativeai ({model_name}): 빈 텍스트 응답이 반환되었습니다.")
        except Exception as e:
            collected_errors.append(f"google-generativeai ({model_name}) 호출 실패: {type(e).__name__} - {str(e)}")
            if model_name != "gemini-3.8-flash":
                try:
                    retry_model = legacy_genai.GenerativeModel("gemini-3.8-flash")
                    retry_res = retry_model.generate_content(prompt)
                    if retry_res and retry_res.text:
                        return retry_res.text
                except Exception as retry_e:
                    collected_errors.append(f"google-generativeai (gemini-3.8-flash 재시도) 실패: {type(retry_e).__name__} - {str(retry_e)}")
    except ImportError:
        collected_errors.append("google-generativeai 패키지가 미설치되어 있습니다.")
    except Exception as e:
        collected_errors.append(f"google-generativeai 클라이언트 초기화 실패: {type(e).__name__} - {str(e)}")

    error_summary = "\n• " + "\n• ".join(collected_errors)
    raise RuntimeError(f"Gemini API 호출에 모두 실패하였습니다. 상세 오류 내용:\n{error_summary}")


def fetch_book_info_and_summary(api_key: str, book_title: str, model_name: str) -> Dict:
    """
    도서명을 기반으로 저자명, 출판 개요 및 챕터/테마별 구조화 요약을 추출합니다.
    """
    prompt = f"""
당신은 기업 사내 독서경영 및 도서 분석 전문가입니다.
사용자가 입력한 도서명 '{book_title}'에 대한 정보를 바탕으로 다음 항목을 JSON 형식으로 작성해 주세요.

반드시 마크다운 코드블록(```json ... ```) 또는 순수 JSON 형식으로 응답하세요.

JSON 필드 규격:
{{
  "author": "저자명 (외국 저자인 경우 한국어 표기 및 원어 병기 가능)",
  "category": "분야/카테고리 (예: 자기계발, 경영전략, 생산/품질, 기술 트렌드 등)",
  "overview": "도서 전체의 핵심 주제 및 전달 가치 요약 (3~4문장)",
  "chapters": [
    {{
      "theme": "핵심 챕터 또는 주제 1 (예: 1장 또는 핵심 테마명)",
      "summary": "해당 챕터/테마의 주요 핵심 내용 및 교훈 요약 (2~3문장)"
    }},
    {{
      "theme": "핵심 챕터 또는 주제 2",
      "summary": "해당 챕터/테마의 주요 핵심 내용 및 교훈 요약 (2~3문장)"
    }},
    {{
      "theme": "핵심 챕터 또는 주제 3",
      "summary": "해당 챕터/테마의 주요 핵심 내용 및 교훈 요약 (2~3문장)"
    }},
    {{
      "theme": "핵심 챕터 또는 주제 4",
      "summary": "해당 챕터/테마의 주요 핵심 내용 및 교훈 요약 (2~3문장)"
    }}
  ]
}}
"""
    raw_response = call_gemini_api(api_key, prompt, model_name)

    # JSON 파싱 시도
    cleaned = raw_response.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)

    try:
        data = json.loads(cleaned)
        return data
    except Exception:
        # 정규식 fallback 추출
        author_match = re.search(r'"author"\s*:\s*"([^"]+)"', raw_response)
        author = author_match.group(1) if author_match else "저자 미상"
        return {
            "author": author,
            "category": "일반",
            "overview": raw_response[:300] + "...",
            "chapters": [
                {"theme": "도서 핵심 내용 요약", "summary": raw_response}
            ],
        }


def generate_idea_sheet_content(
    api_key: str,
    book_title: str,
    author: str,
    department: str,
    user_name: str,
    book_summary: str,
    user_notes: str,
    user_ideas: str,
    model_name: str,
) -> str:
    """
    '7. 내용정리' 칸에 들어갈 독서 요약 및 업무 개선 아이디어를 생성합니다.
    도서 사전 요약(book_summary)이 없거나 공란이어도 도서명과 저자명을 기반으로 AI가 스스로 내용을 유추하여 작성합니다.
    """
    has_summary = bool(book_summary and book_summary.strip())
    has_author = bool(author and author.strip())
    has_user_ideas = bool(user_ideas and user_ideas.strip())
    has_user_notes = bool(user_notes and user_notes.strip())

    # 1. 도서 정보 섹션 구성
    if has_summary:
        book_info_prompt = f"""[도서 정보 및 기확보 요약]
- 도서명: {book_title}
- 저자명: {author if has_author else '저자 미상'}
- 도서 핵심 요약:
{book_summary.strip()}"""
    else:
        book_info_prompt = f"""[도서 정보 (사전 요약 미제공)]
- 도서명: {book_title}
- 저자명: {author if has_author else 'AI 지식으로 자동 유추'}

★ [AI 도서 유추 지침]: 사용자가 사전에 도서 검색/요약을 거치지 않았으므로, Gemini가 보유한 도서 지식을 총동원하여 '{book_title}'(저자: {author if has_author else '해당 도서의 저자'})의 핵심 주제, 출판 배경, 핵심 논지 및 챕터별 주요 시사점을 스스로 파악·유추하여 풍부하게 작성하세요."""

    # 2. 독서 메모 및 개선 아이디어 지침 구성
    instructions = []
    if has_user_ideas:
        instructions.append(f"""- [사용자 개선 아이디어 반영]: 사용자가 직접 제시한 아래 아이디어를 보고서의 2단계(업무 개선안)에 핵심으로 채택하여 현업 적용 방안을 구체화하세요:\n'''{user_ideas.strip()}'''""")
    else:
        instructions.append(f"""- [신입사원 관점 업무 개선 아이디어 독자 도출]: 사용자의 별도 요청 아이디어가 공란이므로, {department} 신입사원({user_name})의 시각에서 도서의 교훈과 핵심 방법론을 현업에 접목할 수 있는 실질적인 업무 개선 아이디어를 최소 2가지 이상 구체적으로 창안하여 제시하세요.
  (예: 현장 공정 낭비 요소 제거, 품질 관리 프로세스 표준화, 5S 및 안전 관리, 디지털/스마트 툴을 활용한 인수인계 효율화, 팀 내 소통 활성화 등 {department}에 실질적 도움이 되는 구체적 방안)""")

    if has_user_notes:
        instructions.append(f"""- [사용자 독서 메모 인용]: 사용자가 독서 중 기록한 아래 메모/인상 깊은 구절을 내용정리의 적절한 위치에 자연스럽게 인용하거나 반영하세요:\n'''{user_notes.strip()}'''""")

    extra_instructions = "\n".join(instructions)

    prompt = f"""당신은 대기업 사내 독서경영 코치이자 비즈니스 문서 작성 전문가입니다.
사내 공식 양식인 [독서경영 IDEA SHEET]의 '7. 내용정리' 란에 기입할 최고 품질의 비즈니스 독서 보고서를 작성해야 합니다.

[작성자 및 부서 정보]
- 부서명: {department}
- 작성자(성명): {user_name}

{book_info_prompt}

{extra_instructions}

[반드시 준수해야 하는 엄격한 제약 사항]
1. 분량 기준 (최우선 필수):
   - 결과 텍스트는 엔터(개행 문자 \\n) 기준 **반드시 15줄 이상의 줄글 또는 문단 구조**로 작성해야 합니다.
   - 각 소주제(1. 도서 핵심 요약, 2. 업무 개선 아이디어, 3. 실천 다짐 및 느낀 점)별로 구체적인 설명과 개행을 풍성하게 두어, 빈 줄을 제외하고도 최소 16~22줄 내외가 되도록 알차게 채우세요.
2. 내용 구성 (3단계 프레임워크 필수 포함):
   - [1단계: 도서 핵심 요약 및 통찰]: 책이 제시하는 핵심 메시지와 주요 원리 (최소 5줄 이상)
   - [2단계: 실무 프로세스 개선 아이디어 제안]: {department} 신입사원 시각에서 현업에 실제 적용할 수 있는 구체적 개선안 및 기대효과 (최소 6줄 이상)
   - [3단계: 신입사원 실천 다짐 및 느낀 점]: 현업 수행 자세, 팀워크, 개인적 성장 목표에 대한 다짐 (최소 4줄 이상)
3. 어조 및 문체:
   - 비즈니스에 적합한 단정하고 논리적이며 성실함이 묻어나는 보고서 문체 (~하였습니다, ~하고자 합니다, ~하겠습니다).
   - 추상적인 미사여구는 지양하고 구체적인 실행 중심의 어조를 유지할 것.
4. 출력 형식:
   - 인사말, 머리말, 꼬리말, 서명 등 불필요한 메타 멘트는 절대 출력하지 말고 오직 PPTX '7. 내용정리' 칸에 그대로 들어갈 본문 텍스트만 출력하세요.
"""
    result = call_gemini_api(api_key, prompt, model_name)
    cleaned_result = result.strip() if result else ""

    # 줄 수 검증 (15줄 미만인 경우 문단 분할 및 보강)
    lines = [line for line in cleaned_result.split("\n") if line.strip()]
    if len(lines) < 15 and cleaned_result:
        refine_prompt = f"""다음은 작성된 독서경영 내용정리 초안입니다.
현재 총 {len(lines)}줄로 작성되어 사내 최소 규격(15줄 이상)에 미달합니다.

내용을 더욱 심도 있고 구체화하여, 각 단락을 적절히 개행(\\n)하고 세부 실행 계획 및 기대효과를 보강하여
**반드시 빈 줄을 제외하고도 16줄 이상 22줄 이하의 풍성한 줄글**로 다시 재작성해 주세요.

[초안]
{cleaned_result}
"""
        try:
            refined = call_gemini_api(api_key, refine_prompt, model_name)
            if refined and refined.strip():
                refined_clean = refined.strip()
                refined_lines = [l for l in refined_clean.split("\n") if l.strip()]
                if len(refined_lines) >= 15:
                    return refined_clean
        except Exception:
            # refine 과정에서 실패하더라도 초안을 반환
            pass

    return cleaned_result


# ==========================================
# 3. PPTX 템플릿 처리 및 폰트 서식 주입 모듈
# ==========================================
def find_target_table(prs: Presentation):
    """
    슬라이드 내에서 '부서명' 또는 '내용정리'가 포함된 독서경영 표를 탐색합니다.
    """
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_table:
                table = shape.table
                # 표 텍스트 전체 합치기
                full_text = "".join(cell.text for row in table.rows for cell in row.cells)
                if "부서명" in full_text or "내용정리" in full_text:
                    return slide, table
    return None, None


def fill_cell_text(
    cell,
    text: str,
    font_size_pt: int = 11,
    bold: bool = False,
    align=PP_ALIGN.CENTER,
    font_name: str = "맑은 고딕",
):
    """
    표 일반 셀에 서식을 유지하며 텍스트를 기입합니다.
    """
    cell.text = ""  # 초기화
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text) if text else ""
    run.font.size = Pt(font_size_pt)
    run.font.bold = bold
    run.font.name = font_name
    run.font.color.rgb = RGBColor(0, 0, 0)


def fill_content_cell(
    cell,
    content_text: str,
    font_size_pt: int = 9,
    font_name: str = "맑은 고딕",
):
    """
    '7. 내용정리' 셀에 9pt 폰트 크기와 문단 줄바꿈을 강제 적용하여 주입합니다.
    """
    cell.text = ""  # 기존 텍스트 초기화
    # 여백 최적화 (텍스트가 넘치지 않도록 세팅)
    cell.margin_left = Inches(0.08)
    cell.margin_right = Inches(0.08)
    cell.margin_top = Inches(0.08)
    cell.margin_bottom = Inches(0.08)

    lines = content_text.split("\n")
    for idx, line in enumerate(lines):
        if idx == 0:
            p = cell.text_frame.paragraphs[0]
        else:
            p = cell.text_frame.add_paragraph()

        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.15
        p.space_after = Pt(2)
        run = p.add_run()
        run.text = line
        run.font.size = Pt(font_size_pt)
        run.font.name = font_name
        run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)


def generate_completed_pptx(
    template_source,
    department: str,
    user_name: str,
    book_title: str,
    author: str,
    start_date: str,
    end_date: str,
    content_text: str,
) -> io.BytesIO:
    """
    지정된 템플릿에 사용자의 입력값과 생성된 내용정리 텍스트를 주입한 뒤 바이너리 스트림으로 반환합니다.
    """
    prs = Presentation(template_source)
    _, table = find_target_table(prs)

    if table is None:
        raise ValueError("PPTX 파일 내에서 독서경영 양식 표를 찾을 수 없습니다. 올바른 템플릿인지 확인해 주세요.")

    # 템플릿 표 구조 매핑
    # Row 0: [0]=부서명(라벨), [1]=부서명값, [2]=성명(라벨), [3]=성명값
    # Row 1: [0]=도서명(라벨), [1]=도서명값, [2]=저자명(라벨), [3]=저자명값
    # Row 2: [0]=시작일(라벨), [1]=시작일값, [2]=완독일(라벨), [3]=완독일값
    # Row 3: [0]=7. 내용정리 라벨/설명문
    # Row 4: [0]=7. 내용정리 본문 기입란 (gridSpan=4 병합 셀)

    fill_cell_text(table.cell(0, 1), department, font_size_pt=11, bold=False, align=PP_ALIGN.CENTER)
    fill_cell_text(table.cell(0, 3), user_name, font_size_pt=11, bold=False, align=PP_ALIGN.CENTER)
    fill_cell_text(table.cell(1, 1), book_title, font_size_pt=11, bold=False, align=PP_ALIGN.CENTER)
    fill_cell_text(table.cell(1, 3), author, font_size_pt=11, bold=False, align=PP_ALIGN.CENTER)
    fill_cell_text(table.cell(2, 1), start_date, font_size_pt=11, bold=False, align=PP_ALIGN.CENTER)
    fill_cell_text(table.cell(2, 3), end_date, font_size_pt=11, bold=False, align=PP_ALIGN.CENTER)

    # 7. 내용정리 칸 (Row 4, Col 0 - 9pt 강제 및 줄바꿈 유지)
    fill_content_cell(table.cell(4, 0), content_text, font_size_pt=9)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output


# ==========================================
# 4. Streamlit 세션 상태 초기화
# ==========================================
DEFAULT_TEMPLATE_PATH = "신입사원_독서경영_양식 (1주차_대경생산팀_도서명).pptx"

if "book_info" not in st.session_state:
    st.session_state.book_info = None
if "content_summary" not in st.session_state:
    st.session_state.content_summary = ""
if "book_title_input" not in st.session_state:
    st.session_state.book_title_input = ""
if "author_input" not in st.session_state:
    st.session_state.author_input = ""
if "start_date_str" not in st.session_state:
    st.session_state.start_date_str = "2026.10.13"
if "end_date_str" not in st.session_state:
    st.session_state.end_date_str = "2026.10.16"
if "final_content_editor" not in st.session_state:
    st.session_state.final_content_editor = ""


# ==========================================
# 5. 사이드바 구성 (설정 및 템플릿 관리)
# ==========================================
with st.sidebar:
    st.title("⚙️ 설정 및 환경")

    # API 키 처리
    env_api_key = os.getenv("GEMINI_API_KEY", "")
    api_key = st.text_input(
        "Google Gemini API Key",
        value=env_api_key,
        type="password",
        help="Google AI Studio(https://aistudio.google.com/)에서 발급받은 API 키를 입력하세요.",
    )
    if not api_key:
        st.warning("⚠️ Gemini API 키를 입력하거나 .env 파일에 등록해 주세요.")
    else:
        st.success("🔑 API 키 준비 완료")

    # 모델 선택
    model_choice = st.selectbox(
        "사용 AI 모델",
        ["gemini-3.8-flash"],
        index=0,
        help="기본으로 속도와 성능이 뛰어난 gemini-3.8-flash가 권장됩니다.",
    )

    st.markdown("---")
    st.subheader("📄 PPTX 템플릿 관리")

    # 기본 템플릿 존재 여부 확인
    has_default_template = os.path.exists(DEFAULT_TEMPLATE_PATH)
    if has_default_template:
        st.info(f"✅ 기본 템플릿 로드됨:\n`{DEFAULT_TEMPLATE_PATH}`")
    else:
        st.warning("기본 템플릿 파일을 찾을 수 없습니다. 아래에서 업로드해 주세요.")

    uploaded_template = st.file_uploader(
        "사용자 정의 PPTX 템플릿 업로드 (선택)",
        type=["pptx"],
        help="기본 템플릿 대신 다른 PPTX 양식을 사용하려면 업로드하세요.",
    )

    st.markdown("---")
    st.markdown(
        """
        **💡 이용 가이드**
        1. 도서명을 입력하고 **[조회]**를 누르면 Gemini가 저자 및 챕터별 핵심 요약을 자동 완성합니다.
        2. 기본 정보(성명, 부서, 주차, 날짜)를 확인 및 수정합니다.
        3. 특별한 아이디어나 독서 메모가 있다면 기입합니다.
        4. **[7. 내용정리 생성]**을 누르면 15줄 이상의 비즈니스 보고서가 완성됩니다.
        5. 필요 시 직접 다듬은 후 **[PPTX 다운로드]**를 진행합니다.
        """
    )


# ==========================================
# 6. 메인 화면 UI
# ==========================================
st.markdown('<div class="main-title">📚 신입사원 독서경영 IDEA SHEET 자동 생성기</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">밀리의 서재 등 독서 후, Gemini AI로 챕터 요약과 업무 개선 아이디어를 도출하여 사내 공식 PPTX 서식에 즉시 주입합니다.</div>',
    unsafe_allow_html=True,
)

tab1, tab2 = st.tabs(["🚀 IDEA SHEET 생성 워크스페이스", "📋 양식 및 규칙 가이드"])

with tab1:
    # ----------------------------------------------------
    # STEP 1. 도서 검색 및 요약 탐색 영역
    # ----------------------------------------------------
    st.markdown("### 1️⃣ 도서 검색 및 핵심 요약 추출")
    col_search1, col_search2 = st.columns([4, 1])

    with col_search1:
        searched_book = st.text_input(
            "도서명 입력",
            value=st.session_state.book_title_input,
            placeholder="예: 원칙(Principles), 피터 드러커의 최고의 질문, 도커 교과서 등",
            key="book_search_field",
        )
    with col_search2:
        st.write("")
        st.write("")
        search_clicked = st.button("🔍 도서 정보 조회", use_container_width=True)

    if search_clicked:
        if not searched_book.strip():
            st.error("도서명을 먼저 입력해 주세요.")
        elif not api_key:
            st.error("Gemini API 키가 필요합니다. 사이드바에 API 키를 입력해 주세요.")
        else:
            with st.spinner(f"'{searched_book}'의 저자 정보 및 챕터별 핵심 내용을 탐색 중입니다..."):
                try:
                    info = fetch_book_info_and_summary(api_key, searched_book, model_choice)
                    st.session_state.book_info = info
                    st.session_state.book_title_input = searched_book
                    st.session_state.author_input = info.get("author", "")
                    st.success(f"도서 정보 탐색 완료! (저자: {st.session_state.author_input})")
                except Exception as e:
                    import traceback
                    st.error(f"❌ 도서 정보 탐색 중 오류가 발생했습니다: {type(e).__name__} - {str(e)}")
                    with st.expander("🔍 에러 상세 정보(Traceback) 확인"):
                        st.code(traceback.format_exc(), language="python")

    # 도서 정보가 있을 때 챕터별 요약 카드/아코디언 표시
    if st.session_state.book_info:
        info = st.session_state.book_info
        with st.expander("📖 AI 분석 도서 개요 및 챕터별 핵심 내용 확인", expanded=True):
            st.markdown(f"**카테고리**: `{info.get('category', '일반')}` | **저자**: `{info.get('author', '저자 미상')}`")
            st.info(f"**전체 개요**: {info.get('overview', '')}")

            chapters = info.get("chapters", [])
            if chapters:
                st.markdown("**주요 챕터 및 테마 요약:**")
                ch_cols = st.columns(min(len(chapters), 3) if len(chapters) <= 3 else 2)
                for idx, ch in enumerate(chapters):
                    col_target = ch_cols[idx % len(ch_cols)]
                    with col_target:
                        st.markdown(
                            f"""
                            <div class="card-box">
                                <b style="color: #1E40AF;">📌 {ch.get('theme', f'주제 {idx+1}')}</b><br/>
                                <span style="font-size: 0.92rem; color: #374151;">{ch.get('summary', '')}</span>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )

    # 대형 텍스트 영역: 원문 메모/인상 깊은 문장
    user_notes = st.text_area(
        "📝 독서 중 메모 또는 인상 깊은 문장 (선택사항)",
        placeholder="밀리의 서재를 읽으며 기록한 주요 발췌문, 감명받은 구절, 핵심 키워드를 자유롭게 붙여넣으세요. 최종 IDEA SHEET 내용 작성 시 우선 반영됩니다.",
        height=100,
    )

    st.markdown("---")

    # ----------------------------------------------------
    # STEP 2. PPT 기입 정보 폼
    # ----------------------------------------------------
    st.markdown("### 2️⃣ PPT 양식 기입 정보 설정")

    col_meta1, col_meta2, col_meta3 = st.columns(3)
    with col_meta1:
        week_value = st.selectbox("주차 (파일명 생성용)", ["1주차", "2주차", "3주차", "4주차", "5주차", "6주차", "7주차", "8주차"], index=0)
    with col_meta2:
        department_value = st.text_input("부서명", value="대경생산팀")
    with col_meta3:
        user_name_value = st.text_input("성명", value="이민재")

    col_book1, col_book2 = st.columns(2)
    with col_book1:
        final_book_title = st.text_input("도서명", value=st.session_state.book_title_input)
    with col_book2:
        final_author = st.text_input("저자명", value=st.session_state.author_input)

    # 날짜 입력 (달력 선택기 & YYYY.MM.DD 텍스트 상호 연동)
    col_date1, col_date2 = st.columns(2)
    with col_date1:
        st.markdown("**5. 시작일**")
        d_col1, d_col2 = st.columns([1, 1])
        with d_col1:
            try:
                parts = [int(p) for p in st.session_state.start_date_str.split(".")]
                default_start = datetime.date(parts[0], parts[1], parts[2])
            except Exception:
                default_start = datetime.date.today()
            picker_start = st.date_input("시작일 달력 선택", value=default_start, key="picker_start")
        with d_col2:
            start_date_input = st.text_input(
                "시작일 텍스트 (YYYY.MM.DD)",
                value=picker_start.strftime("%Y.%m.%d"),
                key="text_start_date",
            )

    with col_date2:
        st.markdown("**6. 완독일**")
        d_col3, d_col4 = st.columns([1, 1])
        with d_col3:
            try:
                parts = [int(p) for p in st.session_state.end_date_str.split(".")]
                default_end = datetime.date(parts[0], parts[1], parts[2])
            except Exception:
                default_end = datetime.date.today()
            picker_end = st.date_input("완독일 달력 선택", value=default_end, key="picker_end")
        with d_col4:
            end_date_input = st.text_input(
                "완독일 텍스트 (YYYY.MM.DD)",
                value=picker_end.strftime("%Y.%m.%d"),
                key="text_end_date",
            )

    # 사용자 요청 텍스트 상자 (프롬프트 반영용 분기 로직)
    user_ideas_prompt = st.text_area(
        "💡 업무 개선 아이디어 및 특별 요청사항 (선택사항)",
        placeholder="예: 공정 불량률 감소를 위한 5S 활동 연계, 부서 간 인수인계 프로세스 디지털화 등 특별히 강조하고 싶은 업무 개선 아이디어나 서술 스타일(예: 개조식과 서술형 혼용)을 입력하세요.\n* 공란으로 둘 경우, 도서 내용을 바탕으로 신입사원 시각의 실무 개선 방안이 자동 구성됩니다.",
        height=90,
    )

    st.markdown("---")

    # ----------------------------------------------------
    # STEP 3. 7. 내용정리 생성 및 편집
    # ----------------------------------------------------
    st.markdown("### 3️⃣ '7. 내용정리' 생성 및 편집 (15줄 이상, 9pt)")

    gen_col1, gen_col2 = st.columns([2, 5])
    with gen_col1:
        create_content_btn = st.button("✨ 7. 내용정리 AI 자동 생성", type="primary", use_container_width=True)

    if create_content_btn:
        effective_book_title = final_book_title.strip() or searched_book.strip()
        effective_author = final_author.strip()

        if not effective_book_title:
            st.error("❌ 도서명이 입력되지 않았습니다. 1단계 또는 2단계에 도서명을 입력해 주세요.")
        elif not api_key:
            st.error("❌ Gemini API 키가 필요합니다. 사이드바에 API 키를 입력해 주세요.")
        else:
            with st.spinner("15줄 이상의 고품질 비즈니스 독서경영 보고서를 작성 중입니다..."):
                try:
                    summary_context = ""
                    if st.session_state.book_info:
                        info_dict = st.session_state.book_info
                        overview = info_dict.get("overview", "")
                        chapters = info_dict.get("chapters", [])
                        if isinstance(chapters, list):
                            ch_texts = [
                                f"[{c.get('theme', '주제')}]: {c.get('summary', '')}"
                                for c in chapters
                                if isinstance(c, dict)
                            ]
                            summary_context = f"{overview}\n" + "\n".join(ch_texts)
                        else:
                            summary_context = str(overview)

                    generated_text = generate_idea_sheet_content(
                        api_key=api_key,
                        book_title=effective_book_title,
                        author=effective_author,
                        department=department_value,
                        user_name=user_name_value,
                        book_summary=summary_context,
                        user_notes=user_notes,
                        user_ideas=user_ideas_prompt,
                        model_name=model_choice,
                    )

                    if not generated_text or not generated_text.strip():
                        st.error("❌ Gemini API로부터 텍스트가 반환되지 않았습니다. 응답 결과가 비어 있습니다.")
                    else:
                        st.session_state.content_summary = generated_text
                        st.session_state["final_content_editor"] = generated_text
                        st.toast("내용정리 작성이 완료되었습니다!", icon="🎉")
                        st.rerun()
                except Exception as e:
                    import traceback
                    err_type = type(e).__name__
                    err_msg = str(e)
                    err_trace = traceback.format_exc()
                    st.error(f"❌ 내용 생성 중 오류가 발생했습니다: [{err_type}] {err_msg}")
                    with st.expander("🔍 에러 상세 원인 및 Traceback 확인", expanded=True):
                        st.code(err_trace, language="python")

    # 텍스트 에어리어로 사용자가 직접 수정할 수 있게 제공
    edited_content = st.text_area(
        "최종 기입될 내용정리 본문 (수정 가능)",
        value=st.session_state.content_summary,
        height=360,
        help="생성된 문구를 직접 보완하거나 다듬을 수 있습니다. 엔터키(\\n)로 구분된 줄 수가 15줄 이상인지 확인하세요.",
        key="final_content_editor",
    )
    # 사용자가 텍스트 에어리어에서 수정한 경우 세션 상태와 실시간 동기화
    if edited_content != st.session_state.content_summary:
        st.session_state.content_summary = edited_content

    # 줄 수 및 글자 수 검증 배지
    non_empty_lines = [l for l in edited_content.split("\n") if l.strip()]
    line_count = len(non_empty_lines)
    char_count = len(edited_content)

    col_stat1, col_stat2, col_stat3 = st.columns([2, 2, 4])
    with col_stat1:
        if line_count >= 15:
            st.markdown(f'<div class="badge-success">✅ {line_count}줄 작성됨 (15줄 이상 규격 충족)</div>', unsafe_allow_html=True)
        elif line_count > 0:
            st.markdown(f'<div class="badge-warning">⚠️ 현재 {line_count}줄 (규격: 15줄 이상 필요)</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="badge-warning">내용이 비어 있습니다.</div>', unsafe_allow_html=True)

    with col_stat2:
        st.caption(f"공백 포함 총 {char_count:,}자")

    st.markdown("---")

    # ----------------------------------------------------
    # STEP 4. PPTX 파일 생성 및 다운로드
    # ----------------------------------------------------
    st.markdown("### 4️⃣ 완성된 PPTX 다운로드")

    # 다운로드 파일명 형식 정의
    actual_book_title = final_book_title.strip() or searched_book.strip()
    clean_book_name = re.sub(r'[\\/*?:"<>|]', "", actual_book_title) or "도서명"
    output_filename = f"신입사원_독서경영_양식 ({week_value}_{department_value}_{clean_book_name}).pptx"

    template_file_to_use = None
    if uploaded_template is not None:
        template_file_to_use = uploaded_template
    elif os.path.exists(DEFAULT_TEMPLATE_PATH):
        template_file_to_use = DEFAULT_TEMPLATE_PATH

    col_down1, col_down2 = st.columns([2, 4])

    with col_down1:
        if template_file_to_use is None:
            st.error("사용 가능한 PPTX 템플릿이 없습니다. 사이드바에서 템플릿 파일을 업로드해 주세요.")
        elif not edited_content.strip():
            st.info("내용정리를 먼저 생성하거나 작성해 주세요.")
        else:
            try:
                pptx_bytes = generate_completed_pptx(
                    template_source=template_file_to_use,
                    department=department_value,
                    user_name=user_name_value,
                    book_title=final_book_title,
                    author=final_author,
                    start_date=start_date_input,
                    end_date=end_date_input,
                    content_text=edited_content,
                )

                st.download_button(
                    label="📥 완성된 PPTX 다운로드",
                    data=pptx_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"PPTX 생성 중 오류 발생: {str(e)}")

    with col_down2:
        st.caption(f"📌 저장 예정 파일명: `{output_filename}`")
        st.caption("표의 7. 내용정리 칸에는 지정된 9pt 폰트 서식과 문단 줄바꿈이 완벽히 적용되어 출력됩니다.")

with tab2:
    st.markdown("### 📋 독서경영 IDEA SHEET 규격 및 서식 가이드")
    st.markdown(
        """
        - **기본 양식**: `신입사원_독서경영_양식 (1주차_대경생산팀_도서명).pptx`
        - **표(Table) 좌표 구성**:
          | 행/열 | 0열 (Col 0) | 1열 (Col 1) | 2열 (Col 2) | 3열 (Col 3) |
          | :--- | :--- | :--- | :--- | :--- |
          | **0행** | 부서명 (라벨) | **부서명 값** (대경생산팀) | 2. 성 명 (라벨) | **성명 값** (이민재) |
          | **1행** | 3. 도서명 (라벨) | **도서명 값** | 4. 저자명 (라벨) | **저자명 값** |
          | **2행** | 5. 시작일 (라벨) | **시작일 값** (YYYY.MM.DD) | 6. 완독일 (라벨) | **완독일 값** (YYYY.MM.DD) |
          | **3행** | <td colspan="4">7. 내용정리 (9pt, 15줄 이상) 안내 헤더</td> | | | |
          | **4행** | <td colspan="4">**7. 내용정리 본문 (9pt 강제, 15줄 이상)**</td> | | | |

        - **핵심 요구조건**:
          1. **15줄 이상 작성**: 단락과 세부 항목 구분을 통해 15줄 이상의 충분한 분량 확보
          2. **9pt 폰트 서식**: 내용정리 기입란의 글자 크기는 9pt(`Pt(9)`)로 고정하여 표 밖으로 텍스트가 넘치지 않도록 자동 제어
          3. **파일명 자동 생성**: `신입사원_독서경영_양식 ({주차}_{부서명}_{도서명}).pptx`
        """
    )

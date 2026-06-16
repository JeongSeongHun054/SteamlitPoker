# -*- coding: utf-8 -*-
import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# 1. 색상 정의 (Premium Dark Mode Palette)
BG_COLOR = RGBColor(15, 23, 42)        # Slate 900 (매우 어두운 네이비 그레이)
CARD_BG_COLOR = RGBColor(30, 41, 59)   # Slate 800 (카드 배경)
TEXT_WHITE = RGBColor(255, 255, 255)   # 타이틀 흰색
TEXT_MUTED = RGBColor(148, 163, 184)   # Slate 400 (설명/본문 회색)
ACCENT_TEAL = RGBColor(20, 184, 166)   # Teal 500 (핵심 하이라이트)
ACCENT_BLUE = RGBColor(59, 130, 246)   # Blue 500 (서브 하이라이트)

FONT_TITLE = "Malgun Gothic"
FONT_BODY = "Arial"

def create_slide_with_bg(prs):
    """어두운 배경이 적용된 슬라이드 생성"""
    blank_layout = prs.slide_layouts[6] # 빈 레이아웃
    slide = prs.slides.add_slide(blank_layout)
    
    # 배경색 지정
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR
    
    return slide

def add_header(slide, title_text, category_text="PORTFOLIO PRESENTATION"):
    """상단 공통 헤더 추가"""
    # 카테고리 텍스트
    cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.3))
    tf_cat = cat_box.text_frame
    tf_cat.word_wrap = True
    tf_cat.margin_left = tf_cat.margin_right = tf_cat.margin_top = tf_cat.margin_bottom = 0
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category_text.upper()
    p_cat.font.name = FONT_BODY
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = ACCENT_TEAL
    
    # 메인 타이틀
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.7), Inches(0.8))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    tf_title.margin_left = tf_title.margin_right = tf_title.margin_top = tf_title.margin_bottom = 0
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.name = FONT_TITLE
    p_title.font.size = Pt(28)
    p_title.font.bold = True
    p_title.font.color.rgb = TEXT_WHITE

def add_card(slide, left, top, width, height, title_text, text_lines):
    """SaaS 스타일 내용 카드 추가 (오버플로우 방지 튜닝 완료)"""
    # 카드 모양 (둥근 직사각형)
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, 
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = CARD_BG_COLOR
    shape.line.color.rgb = RGBColor(71, 85, 105) # Slate 600 경계선
    shape.line.width = Pt(1)
    
    # 카드 텍스트 프레임 생성 (여백 최소화)
    tb = slide.shapes.add_textbox(
        Inches(left + 0.25), Inches(top + 0.25), 
        Inches(width - 0.5), Inches(height - 0.5)
    )
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    
    # 카드 타이틀
    p_title = tf.paragraphs[0]
    p_title.text = title_text
    p_title.font.name = FONT_TITLE
    p_title.font.size = Pt(16)  # 18pt -> 16pt로 축소해 세로 공간 확보
    p_title.font.bold = True
    p_title.font.color.rgb = ACCENT_TEAL
    p_title.space_after = Pt(10) # 간격 축소
    
    # 카드 본문 라인 추가
    for line in text_lines:
        p = tf.add_paragraph()
        p.text = line
        p.font.name = FONT_TITLE
        p.font.size = Pt(11)  # 13pt -> 11pt로 최적화 (텍스트 잘림 현상 방지)
        p.font.color.rgb = TEXT_MUTED
        p.space_after = Pt(4) # 간격 8pt -> 4pt로 축소
        
        # 특정 키워드가 포함될 경우 폰트 색상 및 굵기 하이라이트
        if "★" in line or "**" in line:
            clean_text = line.replace("★", "").replace("**", "")
            p.text = clean_text
            p.font.bold = True
            p.font.color.rgb = TEXT_WHITE
        elif "핵심 역할" in line or "분석 기법" in line or "성과" in line or "의사결정" in line:
            p.font.bold = True
            p.font.color.rgb = ACCENT_BLUE

def add_image_or_placeholder(slide, image_name, left, top, width, height):
    """지정된 이미지를 배치하고, 없을 경우 대안 플레이스홀더 카드 렌더링"""
    brain_dir = r"C:\Users\gnsl1\.gemini\antigravity-ide\brain\88edeb10-dad7-4721-93ed-5bdbacbab2f0"
    image_path = os.path.join(brain_dir, image_name)
    if os.path.exists(image_path):
        try:
            slide.shapes.add_picture(image_path, Inches(left), Inches(top), width=Inches(width), height=Inches(height))
            return
        except Exception as e:
            print(f"Failed to add image {image_name}: {e}")
            
    # Fallback placeholder card (이미지가 없을 경우)
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, 
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = CARD_BG_COLOR
    shape.line.color.rgb = ACCENT_TEAL
    shape.line.width = Pt(1.5)
    
    tb = slide.shapes.add_textbox(Inches(left + 0.3), Inches(top + height/2 - 0.5), Inches(width - 0.6), Inches(1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = f"[Streamlit Dashboard Image: {image_name}]"
    p.alignment = PP_ALIGN.CENTER
    p.font.name = FONT_BODY
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = ACCENT_TEAL

def set_presenter_notes(slide, notes_text):
    """슬라이드 노트(발표자 폰 메모) 설정"""
    notes_slide = slide.notes_slide
    text_frame = notes_slide.notes_text_frame
    text_frame.text = notes_text

def main():
    prs = Presentation()
    # 16:9 와이드스크린 치수 설정 (13.33인치 x 7.5인치)
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    # 공통 레이아웃 좌표 설정 (세로폭 확장: top=1.5, height=5.2)
    CARD_LEFT = 0.8
    CARD_TOP = 1.5
    CARD_WIDTH = 5.6
    CARD_HEIGHT = 5.2
    
    IMG_LEFT = 6.9
    IMG_TOP = 1.5
    IMG_WIDTH = 5.6
    IMG_HEIGHT = 4.2
    
    # =========================================================================
    # SLIDE 1: Title Slide (표지)
    # =========================================================================
    slide1 = create_slide_with_bg(prs)
    
    # 좌측 장식 바 (Teal)
    accent_bar = slide1.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.8), Inches(0.15), Inches(3.8)
    )
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = ACCENT_TEAL
    accent_bar.line.fill.background()
    
    # 타이틀 텍스트 박스
    tb = slide1.shapes.add_textbox(Inches(1.2), Inches(1.7), Inches(11.0), Inches(4.0))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    
    # 카테고리 레이블
    p0 = tf.paragraphs[0]
    p0.text = "DATA ANALYTICS & DECISION PORTFOLIO"
    p0.font.name = FONT_BODY
    p0.font.size = Pt(12)
    p0.font.bold = True
    p0.font.color.rgb = ACCENT_TEAL
    p0.space_after = Pt(20)
    
    # 메인 타이틀
    p1 = tf.add_paragraph()
    p1.text = "비정형 iGaming 로그 분석을 통한\n플레이어 성향 분류 및 의사결정 시뮬레이터 구축"
    p1.font.name = FONT_TITLE
    p1.font.size = Pt(40)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_WHITE
    p1.space_after = Pt(14)
    
    # 서브타이틀
    p2 = tf.add_paragraph()
    p2.text = "비정형 데이터 파이프라인(ETL)부터 통계 분석, 머신러닝, 실시간 대시보드 웹 배포까지"
    p2.font.name = FONT_TITLE
    p2.font.size = Pt(18)
    p2.font.color.rgb = TEXT_MUTED
    p2.space_after = Pt(40)
    
    # 메타정보
    p3 = tf.add_paragraph()
    p3.text = "SK Planet 데이터 직무 PT 면접 발표자료 (5~10분 분량)\n발표자: 지원자"
    p3.font.name = FONT_TITLE
    p3.font.size = Pt(13)
    p3.font.color.rgb = ACCENT_BLUE
    p3.font.bold = True
    
    # 슬라이드 1 노트 설정
    set_presenter_notes(slide1, """[슬라이드 1 발표 대본]
"안녕하십니까, SK Planet 데이터 직무 지원자 [지원자 이름]입니다.
저는 오늘 '비정형 iGaming 로그 분석을 통한 플레이어 성향 분류 및 의사결정 시뮬레이터 구축' 프로젝트를 주제로 발표를 진행하겠습니다.

본 프로젝트는 게임 시스템이 남긴 정제되지 않은 문자열 로그에서 출발하여, 핵심 지표를 계산하고, 통계 분석과 머신러닝을 적용해 유저의 의사결정을 돕는 웹 플랫폼까지 완성해 낸 End-to-End 데이터 파이프라인 프로젝트입니다. 발표를 시작하겠습니다."

[용어 리마인더]
- 비정형 데이터: 컴퓨터가 바로 계산할 수 없는 자유 줄글 형식의 로그 텍스트 파일.""")
    
    # =========================================================================
    # SLIDE 2: Background & Objectives (배경 및 목적)
    # =========================================================================
    slide2 = create_slide_with_bg(prs)
    add_header(slide2, "01. 프로젝트 배경 및 의사결정 목적", "BACKGROUND & OBJECTIVES")
    
    add_card(
        slide2, CARD_LEFT, CARD_TOP, CARD_WIDTH, CARD_HEIGHT, 
        "비즈니스 분석 및 의사결정 목적",
        [
            "■ 분석 배경 및 문제 정의",
            "  - iGaming 플랫폼의 핵심 과제: 유저 파산 및 이탈(Churn) 방어",
            "  - 올바른 플레이 가이드 부재로 인한 급격한 고객 이탈 직면",
            "",
            "■ 지표 분석 및 발견점 (Insight)",
            "  - VPIP, PFR, AF 지표 마트 구축을 통한 다차원 현황 분석",
            "  - 소극적 유저(Loose-Passive)가 가장 심각한 수준인",
            "    평균 -16,393 칩스 적자를 겪으며 파산/이탈함을 규명",
            "",
            "■ ★데이터 기반 최종 의사결정 (Core Decision)",
            "  - 유저 이탈을 방지하고 전략 개선을 유도하기 위해,",
            "  - **'본인 지표 변경 시 미래의 자산 파산 추이를 모의실험해볼 수 있는 실시간 시뮬레이션(Sandbox) 기능을 대시보드에 신설 및 상용 배포 결정'**함."
        ]
    )
    
    # 우측 Streamlit 대시보드 화면 연동
    add_image_or_placeholder(slide2, "tab1_overview_1781002564252.png", IMG_LEFT, IMG_TOP, IMG_WIDTH, IMG_HEIGHT)
    
    # 슬라이드 2 노트 설정
    set_presenter_notes(slide2, """[슬라이드 2 발표 대본]
"첫 번째로 프로젝트의 배경과 의사결정 목적에 대해 말씀드리겠습니다.

포커나 카지노 같은 베팅 플랫폼 서비스에서 가장 해결해야 할 심각한 문제는 '유저의 이탈, 즉 Churn'입니다. 

우측 화면의 '핵심 지표 요약 대시보드' 렌더링 화면을 보시면 알 수 있듯이, 저는 플레이어들의 베팅 지표인 VPIP, PFR, AF 지표 마트를 구축하고 다차원 분석을 진행했습니다. 분석 결과, 적극적으로 베팅하지 않고 소심하게 남의 베팅만 콜하고 따라간 'Loose-Passive' 유저 집단이 가장 심각한 수준인 평균 -16,393 칩의 적자를 내며 파산해 이탈한다는 구체적 사실을 밝혀냈습니다.

이 데이터를 근거로, 저는 첫 번째 '핵심 의사결정'을 내렸습니다. 단순히 '소심하게 베팅하지 마십시오'라는 글로 된 가이드를 주는 데 그치지 않고, 유저가 직접 마우스로 수치를 조정하며 스스로의 미래 자산 추이를 확인하고 습관을 교정할 수 있는 '실시간 샌드박스 시뮬레이터 기능을 대시보드 플랫폼에 핵심 스펙으로 포함하여 상용 배포하기로 결정'한 것입니다."

[용어 리마인더]
- Churn(이탈): 포커 게임에서는 올바른 습관을 가지지 못해 파산하고 아예 앱을 삭제하는 최악의 유저 행동.
- VPIP (자발적 참여율): 본인의 의지로 돈을 걸고 게임에 참여한 게임 판수의 비율.""")

    # =========================================================================
    # SLIDE 3: Data Specification & Pipeline (데이터와 역할)
    # =========================================================================
    slide3 = create_slide_with_bg(prs)
    add_header(slide3, "02. 활용 데이터 및 ETL 파이프라인 수립", "DATA & ETL PIPELINE")
    
    add_card(
        slide3, CARD_LEFT, CARD_TOP, CARD_WIDTH, CARD_HEIGHT, 
        "비정형 데이터 모델링 및 적재 역할",
        [
            "■ 활용 데이터 (What)",
            "  - 140개 비정형 텍스트 로그 파일 (Hand History)",
            "  - ★총 7,993개 핸드 (Games), 932명 플레이어 데이터",
            "  - ★10만 행 이상의 플레이어별 상세 액션 로그 분석",
            "",
            "■ 본인의 핵심 설계 역할 (How)",
            "  - **Regex 파서 설계**: 텍스트 형태의 원천 로그에서",
            "    베팅액, 플레이어명, 액션 단계를 필터 오차 없이 파싱",
            "  - **스타 스키마 모델링**: DW 설계 표준인 Fact(`hands`,",
            "    `actions`) 및 Dimension(`players`, `tournaments`) 구조 설계",
            "  - **이중 DB 호환**: SQLite 로컬 DB와 PostgreSQL 상용 DB 간에",
            "    환경 변수 하나로 스위칭 전환 가능한 호환 커넥터 구현"
        ]
    )
    
    # 우측 Streamlit 데이터 마트 화면 연동
    add_image_or_placeholder(slide3, "tab2_profiling_1781002574421.png", IMG_LEFT, IMG_TOP, IMG_WIDTH, IMG_HEIGHT)
    
    # 슬라이드 3 노트 설정
    set_presenter_notes(slide3, """[슬라이드 3 발표 대본]
"두 번째로 활용한 데이터 스펙과 데이터 파이프라인 구축 과정에서 제가 수행한 역할입니다.

원천 데이터는 140개의 비정형 포커 게임 텍스트 로그 파일입니다. 약 7,993개의 게임 판수와 932명의 플레이어가 남긴 10만 행 이상의 상세 배팅 로그로 구성되어 있었습니다.

저는 이 비정형 데이터를 컴퓨터가 계산 가능한 관계형 데이터베이스로 구축하기 위해 세 가지를 직접 설계했습니다.

첫째, Python 정규표현식(Regex) 엔진을 자체 설계하여 텍스트 줄글에서 칼같이 필요한 텍스트 정보만 발라내고 오차 없이 파싱하였습니다.
둘째, 우측에 보이는 '플레이어 프로파일링 데이터 마트' 화면을 뒷받침하기 위해 스타 스키마(Star Schema) 모델링을 설계했습니다. 뼈대가 되는 Fact 테이블과 주변 설명 Dimension 테이블을 도출해 DB에 적재했습니다.
셋째, 환경 변수만 변경하면 분석용 SQLite와 상용 PostgreSQL DB로 즉각 적재 경로가 바뀌는 이중 DB 호환 구조를 구현했습니다. 이 데이터 아키텍처 덕분에 쿼리 연산 속도를 크게 높여 대화형 대시보드 구성이 가능해졌습니다."

[용어 리마인더]
- 정규표현식 (Regex): 특정 글자 패턴만 정확하게 추출하는 컴퓨터용 필터 규칙.
- 스타 스키마 (Star Schema): 마트 진열처럼 핵심 기록(Fact)을 중심에 두고 주변에 설명 정보(Dimension)를 별 모양처럼 연결해 검색 효율을 높인 것.""")

    # =========================================================================
    # SLIDE 4: Core Statistics & ML (가설 검정 및 모델링)
    # =========================================================================
    slide4 = create_slide_with_bg(prs)
    add_header(slide4, "03. 통계적 가설 검정 및 의사결정 프로세스", "STATISTICS & MACHINE LEARNING")
    
    add_card(
        slide4, CARD_LEFT, CARD_TOP, CARD_WIDTH, CARD_HEIGHT, 
        "통계 기반 가설 검정 및 성향 예측 모델",
        [
            "■ 가설 설정 및 A/B 테스트 검정",
            "  - 가설: 프리미엄 카드 진입 시 선공(A)이 소극적 콜(B)보다 수익이 높을 것이다.",
            "  - 독립표본 T-검정 수행 -> p-value = 0.6372로 가설 기각.",
            "",
            "■ ★데이터 분석 기반의 프로세스 의사결정",
            "  - 대조군(소극 진입)의 표본수가 극히 작아 생긴 통계 오류(검정력 부족) 진단",
            "  - **'통계적 신뢰성을 가질 때까지 성급한 비즈니스 지침 판단을 유보하고, 성향점수 매칭(PSM) 및 부트스트랩을 도입하여 통계 엔진을 전면 보완하기로 의사결정'**함.",
            "",
            "■ 군집화 및 예측 모델",
            "  - K-Means 알고리즘 적용: 유저 4대 성향 코호트 자동 분류",
            "  - Random Forest Classifier 연동: 유저 흑자 여부 예측 (Accuracy 84%)"
        ]
    )
    
    # 우측 Streamlit 가설 검정 화면 연동
    add_image_or_placeholder(slide4, "tab4_ab_test_1781002593463.png", IMG_LEFT, IMG_TOP, IMG_WIDTH, IMG_HEIGHT)
    
    # 슬라이드 4 노트 설정
    set_presenter_notes(slide4, """[슬라이드 4 발표 대본]
"세 번째로 통계적 가설 검정과 머신러닝 분석 과정에서의 데이터 기반 의사결정 경험입니다.

저는 '좋은 패를 들었을 때 소극적으로 콜만 하고 들어가는 유저보다, 먼저 판돈을 올리며(Raise) 적극적으로 선제공격하는 유저가 진짜 돈을 잘 벌 것인가'를 확인하고자 독립표본 T-검정을 설계했습니다. 우측의 'A/B 테스트 가설 검정' 화면을 보시면 그 결과가 나타나 있습니다.
검정 결과 p-value가 0.6372로 나와 통계적으로 유의미한 차이가 없음으로 가설이 기각되었습니다.

여기서 저의 중요한 분석적 의사결정이 발생했습니다. 단순히 통계 수치만 믿고 '두 전략은 차이가 없습니다'라고 플랫폼 유저에게 알렸다면 그것은 잘못된 행동 가이드였을 것입니다. 저는 심층 분석을 진행해, 대조군인 콜 진입 그룹의 표본 수가 극도로 작아 통계적 검정의 오류 가능성이 매우 높음을 찾아냈습니다.
이에 따라 저는 '통계적 데이터가 충분히 보완될 때까지 섣부른 비즈니스 판단과 가이드를 유보하고, 성향점수 매칭(PSM)과 부트스트랩 리샘플링 기술을 데이터 가공 엔진에 즉시 주입하도록 설계 프로세스를 보완하는 의사결정'을 내렸습니다.

또한, K-Means 군집 분석을 통해 플레이어들을 4대 성향군으로 분류하고, 최종적으로 플레이어의 지표를 넣어 이 플레이어가 흑자일지 적자일지 가려내는 Random Forest 분류 모델을 연동하여 정확도 84% 성능을 획득했습니다."

[용어 리마인더]
- T-검정 (A/B Test): 두 그룹의 평균 차이가 진짜 의미 있는 실력 차이인지, 우연인지 검증하는 도구.
- p-value (유의확률): 우연히 그렇게 다를 확률. 0.05 미만이어야 진짜 실력 차이로 인정.
- Random Forest: 스무고개 놀이를 하듯 다수결의 질문 트리를 모아 최종 판단(돈 딸지 잃을지)을 내리는 모델.""")

    # =========================================================================
    # SLIDE 5: Actionable Outputs (분석 결과의 서비스화)
    # =========================================================================
    slide5 = create_slide_with_bg(prs)
    add_header(slide5, "04. 분석 결과의 서비스화 및 의사결정 시뮬레이터", "PLATFORM SERVICE & SIMULATOR")
    
    add_card(
        slide5, CARD_LEFT, CARD_TOP, CARD_WIDTH, CARD_HEIGHT, 
        "시뮬레이터 구현 및 유저 행동 변화 유도",
        [
            "■ 의사결정 보조 플랫폼 구현 (Streamlit)",
            "  - Tableau 스타일 필터: 사이드바 필터 변경 즉시 SQL 동적 렌더링",
            "  - 실시간 ETL 데모: 업로드 시 파서 작동 및 실시간 진행바 표시",
            "",
            "■ ★플레이 스타일 샌드박스 시뮬레이터",
            "  - 유클리드 거리 기준으로 DB 내 나와 지표가 가장 근접한 실제 유저 3명 자동 추적 매칭",
            "  - **의사결정 활용**: VPIP, PFR, AF 조정 시 매칭 유저의 30세션 누적 자산 추이 대조 렌더링",
            "  - **비즈니스 임팩트**: 무리한 베팅(Loose-Passive) 습관이 미래 자산 파산에 미치는 영향을 모의실험하게 유도해 유저 행동 개선 자극"
        ]
    )
    
    # 우측 Streamlit 시뮬레이터 화면 연동
    add_image_or_placeholder(slide5, "tab5_simulator_done_1781002631767.png", IMG_LEFT, IMG_TOP, IMG_WIDTH, IMG_HEIGHT)
    
    # 슬라이드 5 노트 설정
    set_presenter_notes(slide5, """[슬라이드 5 발표 대본]
"네 번째로 분석 결과를 실제 서비스로 전환하여 유저의 의사결정을 도운 부분입니다.

데이터 분석 결과인 최적의 플레이 스타일 범위를 유저에게 가르치고 실제로 플레이 습관을 고치게 하려면, 정적인 글보다는 시각적인 모의실험 도구가 필요했습니다.

우측에 보이시는 '시나리오 시뮬레이터' 화면이 바로 그것입니다. 이 플랫폼의 샌드박스 플레이그라운드에서 유저가 본인의 VPIP, PFR, AF 슬라이더를 마우스로 자유롭게 조정하면, 머신러닝의 유클리드 거리 공식을 통해 데이터베이스 내 932명의 유저 중 지표상 가장 비슷한 성향을 가진 실제 유저 3명을 매칭합니다.
그리고 그 매칭된 유저 3명이 30세션 동안 게임을 치르면서 누적시킨 칩의 시계열 자산 추이 그래프를 대조하여 즉각 그려줍니다.

만약 유저가 돈을 크게 잃고 이탈할 수밖에 없는 Loose-Passive 상태를 설정하면, 화면에는 3인의 유저 자산 그래프가 가파르게 우하향하며 파산하는 모습을 직접 목격하게 됩니다. 유저는 이 시뮬레이션을 통해 본인의 잘못된 플레이 전략이 미래 자산에 미칠 충격을 직접 체감하고, 지표를 개선(VPIP 15~22%, AF 2.0~3.5)하여 우상향하는 그래프를 매칭받도록 본인의 플레이 전략을 스스로 변경하는 행동 개선 의사결정을 내릴 수 있게 됩니다."

[용어 리마인더]
- 유클리드 거리: 지표 공간 내 두 점 간의 직선거리. 여기선 '나와 가장 성격이 비슷한 3명의 유저'를 찾을 때 쓰는 거리 계산 공식.
- 샌드박스: 가상의 공간에서 안전하게 자유로운 모의실험을 해볼 수 있는 환경.""")

    # =========================================================================
    # SLIDE 6: Successes & Improvements (성과 및 개선점)
    # =========================================================================
    slide6 = create_slide_with_bg(prs)
    add_header(slide6, "05. 성과 및 향후 고도화 로드맵", "SUCCESS & ROADMAP")
    
    add_card(
        slide6, CARD_LEFT, CARD_TOP, CARD_WIDTH, CARD_HEIGHT, 
        "분석 과정의 성과 및 향후 개선 사항",
        [
            "■ 분석 과정에서 잘된 점 (Successes)",
            "  - **비정형 데이터의 RDBMS 정형화**: 자유 텍스트 로그를 Fact/Dimension 스타 스키마로 설계하여 다차원 글로벌 쿼리 성능 극대화",
            "  - **분석의 서비스화(MaaS) 성공**: 단방향 수치 요약 보고서에 그치지 않고 인터랙티브 시뮬레이션 웹서비스로 상용 배포하여 데이터 활용성 극대화",
            "",
            "■ 개선 필요 사항 및 향후 계획 (Roadmap)",
            "  - **통계적 신뢰도 보완**: T-Test의 대조군 표본 부족 현상 해결을 위해 성향점수 매칭(PSM) 및 부트스트랩 리샘플링 적용 추진",
            "  - **아키텍처 확장**: 단일 RDBMS 동시성 병목 우려 대응을 위한 실시간 Apache Kafka 큐, Apache Spark 분산 엔진 도입 및 GCP BigQuery 클라우드 DW 마이그레이션 로드맵 수립"
        ]
    )
    
    # 우측 Streamlit 퍼널/코호트 분석 화면 연동
    add_image_or_placeholder(slide6, "tab3_funnel_cohort_1781002583397.png", IMG_LEFT, IMG_TOP, IMG_WIDTH, IMG_HEIGHT)
    
    # 슬라이드 6 노트 설정
    set_presenter_notes(slide6, """[슬라이드 6 발표 대본]
"마지막으로 본 프로젝트의 성과와 향후 개선점에 대해 말씀드리겠습니다.

분석 과정에서의 대표적인 성과는 첫째, 텍스트 형태의 비정형 로그를 표준 RDBMS 스타 스키마로 적재하여 다차원 글로벌 조회 쿼리 성능을 대폭 향상하고 데이터웨어하우스 기반을 확립한 점입니다. 둘째, 단방향 통계 분석에 그치지 않고 유저의 실질적인 습관 교정과 의사결정을 돕는 인터랙티브 웹 시뮬레이터 플랫폼을 개발하고 완전히 오픈 상용 배포했다는 점입니다.

반면, 분석 과정에서의 통계적 한계와 아키텍처 확장성에 대한 개선 방향도 명확히 정의했습니다.

첫째, A/B 테스트에서 대조군의 모수가 적어 기각된 문제는 향후 성향점수 매칭(PSM) 기법과 부트스트랩 리샘플링 기술을 활용해 표본 편향을 수학적으로 완전히 극복할 계획입니다.

둘째, 향후 수천만 건 이상의 트래픽이 몰리는 대형 플랫폼 환경을 감안하면 단일 DBMS 아키텍처는 동시성 락으로 인해 병목이 생길 수 있습니다. 이를 고도화하기 위해 실시간 데이터 수집 큐인 Apache Kafka와 대용량 분산 처리를 위한 Apache Spark를 도입하고, 최종 저장소를 구글 클라우드의 BigQuery 데이터웨어하우스로 마이그레이션하는 빅데이터 파이프라인 로드맵을 수립하고 연구하고 있습니다.

데이터의 정제 수집부터 아키텍처 구성, 분석 임팩트 증명 및 최종 웹앱 서비스 배포까지 데이터의 흐름 전체를 주도해 본 경험을 바탕으로, SK Planet에서 실질적인 비즈니스 가치를 만들어 내는 기여자가 되겠습니다. 이상으로 발표를 마치겠습니다. 감사합니다."

[용어 리마인더]
- 성향점수 매칭 (PSM): A/B 테스트에서 두 집단의 조건 불균형을 잡기 위해 수학적으로 조건이 비슷한 짝을 찾아 맞춰주는 기법.
- Kafka & Spark: 실시간으로 쏟아지는 수백만 건의 대규모 트래픽 데이터를 처리해 주는 글로벌 표준 데이터 스택.""")

    # PPT 저장
    ppt_path = "poker_analytics_presentation.pptx"
    prs.save(ppt_path)
    print(f"Presentation saved successfully to: {os.path.abspath(ppt_path)}")

if __name__ == "__main__":
    main()

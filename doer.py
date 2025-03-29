import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR, MSO_FILL

# 白と黒を基調としたシンプルなカラーパレット
class ColorPalette:
    BACKGROUND = RGBColor(255, 255, 255)  # ホワイト
    TEXT = RGBColor(0, 0, 0)              # ブラック
    ACCENT = RGBColor(50, 50, 50)         # ダークグレー
    FOOTER_BG = RGBColor(240, 240, 240)   # ライトグレー
    FOOTER_TEXT = RGBColor(100, 100, 100) # ミディアムグレー
    HEADING_BG = RGBColor(0, 0, 0)        # ブラック
    HEADING_TEXT = RGBColor(255, 255, 255) # ホワイト
    PANEL_BG = RGBColor(255, 255, 255)      # ホワイト
    PANEL_BORDER = RGBColor(220, 220, 220)  # ライトグレーボーダー

# モダンで洗練されたフォント設定
TITLE_FONT = 'Lato'  # Google Fonts でも利用可能なモダンフォント
BODY_FONT = 'Lato'
# ALT_FONT = 'Helvetica Neue'  # 代替はコメントアウト

# フォントサイズ定義 (余白を意識して調整)
TITLE_SIZE = Pt(48)
HEADING_SIZE = Pt(32)
SUBHEADING_SIZE = Pt(20)
BODY_SIZE = Pt(16)
CAPTION_SIZE = Pt(12)

def create_presentation():
    prs = Presentation()
    
    # スライドサイズを16:9に設定
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    # スライドを作成
    create_title_slide(prs)
    create_executive_summary(prs)
    create_current_analysis(prs)
    create_proposal(prs)
    create_schedule(prs)
    create_team_structure(prs)
    create_risk_management(prs)
    create_budget(prs)
    create_success_criteria(prs)
    create_conclusion(prs)
    
    # プレゼンテーションを保存
    prs.save('great1.pptx')
    print("洗練されたプレゼンテーションが作成されました: great1.pptx")

def apply_title_style(title_shape, text, font_size=TITLE_SIZE, color=ColorPalette.TEXT, align=PP_ALIGN.LEFT, bold=True):
    """タイトルのスタイルを適用する"""
    title_shape.text = text
    title_para = title_shape.text_frame.paragraphs[0]
    title_para.alignment = align
    run = title_para.runs[0] if title_para.runs else title_para.add_run()
    run.font.name = TITLE_FONT
    run.font.size = font_size
    run.font.bold = bold
    run.font.color.rgb = color

def apply_body_style(body_shape, text_list, font_size=BODY_SIZE, color=ColorPalette.TEXT, para_spacing=Pt(12)):
    """本文のスタイルを適用する"""
    tf = body_shape.text_frame
    tf.clear()  # テキストフレーム内のすべての段落をクリア
    # 最初の段落を追加
    p = tf.add_paragraph()
    
    for i, original_text in enumerate(text_list):
        text_to_set = original_text  # スタイル適用用のテキスト

        # 2番目以降は新しい段落を追加
        if i > 0:
            p = tf.add_paragraph()

        # 空文字列または空白のみの場合
        if not original_text.strip():
            p.text = ""
            p.space_after = para_spacing
            continue

        # レベル決定とテキスト変更
        level = 0
        space_before = Pt(0)
        is_heading = False
        is_bullet1 = False
        is_bullet2 = False

        if original_text.startswith('【') and original_text.endswith('】'):
            level = 0
            space_before = Pt(15) if i > 0 else Pt(0)
            is_heading = True
        elif original_text.startswith('• '):
            level = 1
            text_to_set = original_text[2:]
            is_bullet1 = True
        elif original_text.startswith('  - '):
            level = 2
            text_to_set = original_text[4:]
            is_bullet2 = True
        else:
            level = 0

        # テキスト、レベル、スペースを設定
        p.text = text_to_set
        p.level = level
        p.space_after = para_spacing
        p.space_before = space_before

        run = p.runs[0] if p.runs else p.add_run()
        run.font.name = BODY_FONT
        run.font.size = font_size

        if is_heading:
            run.font.bold = True
            run.font.size = font_size + Pt(4)  # 見出しは少し大きく
            run.font.color.rgb = color
        else:
            run.font.bold = False
            run.font.color.rgb = color

def add_shape(slide, shape_type, left, top, width, height, fill_color=None, line_color=None, 
              line_width=Pt(0.75), shadow=False, transparency=0, gradient_to=None, text=None):
    """洗練された図形を追加する (影はデフォルトOFF)"""
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    
    # 塗りつぶし設定
    if fill_color:
        if gradient_to:
            # グラデーション (※pptxでのグラデーション設定は環境によって異なる場合あり)
            shape.fill.gradient()
            shape.fill.gradient_stops[0].position = 0
            shape.fill.gradient_stops[0].color.rgb = fill_color
            shape.fill.gradient_stops[1].position = 1
            shape.fill.gradient_stops[1].color.rgb = gradient_to
            shape.fill.gradient_angle = 90  # 左から右へのグラデーション
        else:
            # 単色
            shape.fill.solid()
            shape.fill.fore_color.rgb = fill_color
            shape.fill.transparency = transparency
    else:
        # 塗りつぶしなし
        shape.fill.background()
    
    # 線の設定
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_width
    else:
        shape.line.fill.background()  # 線なし
    
    # 影の設定
    if shadow:
        shape.shadow.inherit = False
        shape.shadow.visible = True
        shape.shadow.blur_radius = Pt(5)
        shape.shadow.distance = Pt(3)
        shape.shadow.angle = 45
        try:
            shape.shadow.color.rgb = RGBColor(180, 180, 180)  # 薄いグレーの影
            shape.shadow.transparency = 0.5
        except AttributeError:
            pass
    
    # テキストがある場合
    if text:
        shape.text = text
        tf = shape.text_frame
        tf.word_wrap = True
        tf.auto_size = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.runs[0] if p.runs else p.add_run()
        run.font.name = TITLE_FONT
        run.font.size = SUBHEADING_SIZE
        run.font.color.rgb = ColorPalette.HEADING_TEXT
    
    return shape

def add_background(slide, prs, type="solid", color=ColorPalette.BACKGROUND, gradient_to=None):
    """スライドの背景を設定する (シンプルに単色のみ)"""
    background = add_shape(
        slide, MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0),
        prs.slide_width, prs.slide_height,
        fill_color=color,
        line_color=None  # 背景に線は不要
    )
    return background

def add_footer(slide, prs, text="Your Company Name | Project Proposal | 2025", include_page_number=True, page_num=None):
    """フッターを追加する (シンプルデザイン)"""
    footer_shape = add_shape(
        slide, MSO_SHAPE.RECTANGLE,
        Inches(0), prs.slide_height - Inches(0.3),
        prs.slide_width, Inches(0.3),
        fill_color=ColorPalette.FOOTER_BG,
        line_color=None
    )
    
    footer_text = slide.shapes.add_textbox(
        Inches(0.5), prs.slide_height - Inches(0.35),
        prs.slide_width - Inches(1.5), Inches(0.3)
    )
    tf = footer_text.text_frame
    p = tf.paragraphs[0]
    if include_page_number and page_num:
        p.text = f"{text} | {page_num}"
    else:
        p.text = text
    p.alignment = PP_ALIGN.LEFT
    run = p.runs[0] if p.runs else p.add_run()
    run.font.name = BODY_FONT
    run.font.size = CAPTION_SIZE
    run.font.color.rgb = ColorPalette.FOOTER_TEXT
    
    return footer_shape

def create_title_slide(prs):
    """洗練された表紙スライドの作成 (ミニマルデザイン)"""
    slide_layout = prs.slide_layouts[5]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # 背景色
    add_background(slide, prs)

    # タイトル
    title_box = slide.shapes.add_textbox(
        Inches(1), Inches(2.5),
        Inches(11), Inches(2)
    )
    title_tf = title_box.text_frame
    # テキスト設定で run が自動生成されるため直接設定
    title_tf.text = "IT Development & System Implementation"
    title_p = title_tf.paragraphs[0]
    title_p.alignment = PP_ALIGN.LEFT
    run = title_p.runs[0] if title_p.runs else title_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = TITLE_SIZE
    run.font.bold = True
    run.font.color.rgb = ColorPalette.TEXT

    # サブタイトル
    subtitle_p = title_tf.add_paragraph()
    subtitle_p.text = "Project Proposal"
    subtitle_p.alignment = PP_ALIGN.LEFT
    run = subtitle_p.runs[0] if subtitle_p.runs else subtitle_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = SUBHEADING_SIZE
    run.font.bold = False
    run.font.color.rgb = ColorPalette.TEXT

    # 日付と会社名（フッター用）
    details_box = slide.shapes.add_textbox(
        Inches(1), Inches(5.5),
        Inches(11), Inches(0.5)
    )
    details_tf = details_box.text_frame
    details_tf.text = "March 30, 2025 | Your Company Name"
    details_p = details_tf.paragraphs[0]
    details_p.alignment = PP_ALIGN.LEFT
    run = details_p.runs[0] if details_p.runs else details_p.add_run()
    run.font.name = BODY_FONT
    run.font.size = BODY_SIZE
    run.font.color.rgb = ColorPalette.FOOTER_TEXT

    # フッター
    add_footer(slide, prs, text="Your Company Name", page_num="1/10")

def create_executive_summary(prs):
    """洗練されたエグゼクティブサマリーのスライド"""
    slide_layout = prs.slide_layouts[5]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    add_background(slide, prs)
    
    # ヘッダーバー
    add_shape(
        slide, MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0),
        prs.slide_width, Inches(1),
        fill_color=ColorPalette.HEADING_BG
    )
    
    # ヘッダータイトル
    header_title = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.2),
        Inches(12), Inches(0.6)
    )
    header_tf = header_title.text_frame
    header_tf.text = "Executive Summary"
    header_p = header_tf.paragraphs[0]
    header_p.alignment = PP_ALIGN.LEFT
    run = header_p.runs[0] if header_p.runs else header_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = HEADING_SIZE
    run.font.bold = True
    run.font.color.rgb = ColorPalette.HEADING_TEXT

    # コンテンツエリア
    content_box = slide.shapes.add_textbox(
        Inches(1), Inches(1.5),
        Inches(11), Inches(5)
    )
    
    summary_points = [
        "【Project Objective】",
        "• Revamp the current business system to improve operational efficiency by 30%.",
        "• Build a foundation for digital transformation.",
        "",
        "【Key Proposal】",
        "• Implement a cloud-based integrated management system.",
        "• Utilize AI for business process automation and predictive analytics.",
        "",
        "【Expected Benefits】",
        "• Achieve annual cost savings of ¥20 million and reduce customer response time by 50%.",
        "• Enable data-driven decision-making and expand business opportunities."
    ]
    apply_body_style(content_box, summary_points, para_spacing=Pt(14))

    # 予算・期間の情報
    info_box = slide.shapes.add_textbox(
        Inches(8), Inches(5.5),
        Inches(4.5), Inches(1.2)
    )
    info_points = [
        "• Duration: 6 months (Apr 2025 - Sep 2025)",
        "• Budget: ¥35M initial, ¥8M annual running cost",
        "• ROI: Estimated payback in 18 months"
    ]
    tf = info_box.text_frame
    tf.clear()
    p = tf.add_paragraph()
    for text in info_points:
        p.text = text
        p.space_after = Pt(6)
        run = p.runs[0] if p.runs else p.add_run()
        run.font.name = BODY_FONT
        run.font.size = Pt(14)
        run.font.color.rgb = ColorPalette.FOOTER_TEXT
        p = tf.add_paragraph()

    add_footer(slide, prs, include_page_number=True, page_num="2/10")

def create_current_analysis(prs):
    """洗練された現状分析と課題のスライド"""
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    
    add_background(slide, prs)
    
    # ヘッダーバー
    add_shape(
        slide, MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0),
        prs.slide_width, Inches(1),
        fill_color=ColorPalette.HEADING_BG
    )
    
    # ヘッダータイトル
    header_title = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.2),
        Inches(12), Inches(0.6)
    )
    header_tf = header_title.text_frame
    header_tf.text = "Current Situation & Challenges"
    header_p = header_tf.paragraphs[0]
    header_p.alignment = PP_ALIGN.LEFT
    run = header_p.runs[0] if header_p.runs else header_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = HEADING_SIZE
    run.font.bold = True
    run.font.color.rgb = ColorPalette.HEADING_TEXT

    # 左側：現状
    current_box = slide.shapes.add_textbox(
        Inches(1), Inches(1.5),
        Inches(5.5), Inches(5)
    )
    current_tf = current_box.text_frame
    current_tf.text = "Current System Situation"
    current_p = current_tf.paragraphs[0]
    current_p.space_after = Pt(12)
    run = current_p.runs[0] if current_p.runs else current_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = SUBHEADING_SIZE
    run.font.bold = True
    run.font.color.rgb = ColorPalette.TEXT

    current_state = [
        "• Core system operational for 8 years.",
        "• Multiple systems lack integration, causing duplicate data entry.",
        "• Increased maintenance costs due to legacy systems.",
        "• Resource constraints in the on-premises environment.",
        "• Lack of mobile support restricts remote work."
    ]
    apply_body_style(current_box, current_state, font_size=BODY_SIZE, para_spacing=Pt(12))
    # 最初の段落（サブヘッダー）を空にする
    current_tf.paragraphs[0].text = ""

    # 右側：課題
    challenge_box = slide.shapes.add_textbox(
        Inches(7), Inches(1.5),
        Inches(5.5), Inches(5)
    )
    challenge_tf = challenge_box.text_frame
    challenge_tf.text = "Key Challenges to Address"
    challenge_p = challenge_tf.paragraphs[0]
    challenge_p.space_after = Pt(12)
    run = challenge_p.runs[0] if challenge_p.runs else challenge_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = SUBHEADING_SIZE
    run.font.bold = True
    run.font.color.rgb = ColorPalette.TEXT

    challenges = [
        "• Centralize data management and standardize business processes.",
        "• Eliminate redundant work through automated system integration.",
        "• Optimize costs by migrating to a cloud environment.",
        "• Establish a remote work environment with mobile support.",
        "• Enhance security and ensure compliance."
    ]
    apply_body_style(challenge_box, challenges, font_size=BODY_SIZE, para_spacing=Pt(12))
    challenge_tf.paragraphs[0].text = ""

    add_footer(slide, prs, include_page_number=True, page_num="3/10")

def create_proposal(prs):
    """提案内容のスライド (ミニマル)"""
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    
    add_background(slide, prs)
    
    # ヘッダーバー
    add_shape(
        slide, MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0),
        prs.slide_width, Inches(1),
        fill_color=ColorPalette.HEADING_BG
    )
    
    # ヘッダータイトル
    header_title = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.2),
        Inches(12), Inches(0.6)
    )
    header_tf = header_title.text_frame
    header_tf.text = "Proposal: Cloud Integrated Management System"
    header_p = header_tf.paragraphs[0]
    header_p.alignment = PP_ALIGN.LEFT
    run = header_p.runs[0] if header_p.runs else header_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = HEADING_SIZE
    run.font.bold = True
    run.font.color.rgb = ColorPalette.HEADING_TEXT

    # メインコンテンツエリア
    content_box = slide.shapes.add_textbox(
        Inches(1), Inches(1.5),
        Inches(11), Inches(5.5)
    )
    
    main_content = [
        "【System Features】",
        "• Centralized management of all business data.",
        "• Cloud-based platform accessible from anywhere.",
        "• Intuitive user interface.",
        "• Real-time data synchronization and analysis.",
        "• Efficiency gains through business process automation.",
        "• AI-powered predictive analytics and decision support.",
        "• Flexible scalability and customization.",
        "",
        "【Key Functions】",
        "• Integrated customer information and case management.",
        "• Real-time dashboards.",
        "• Workflow automation.",
        "• Role-based access control and security.",
        "• Mobile application support."
    ]
    apply_body_style(content_box, main_content, para_spacing=Pt(14))

    add_footer(slide, prs, include_page_number=True, page_num="4/10")

def create_schedule(prs):
    """導入スケジュールのスライド (テキストベース)"""
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    
    add_background(slide, prs)
    
    # ヘッダーバー
    add_shape(
        slide, MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0),
        prs.slide_width, Inches(1),
        fill_color=ColorPalette.HEADING_BG
    )
    
    # ヘッダータイトル
    header_title = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.2),
        Inches(12), Inches(0.6)
    )
    header_tf = header_title.text_frame
    header_tf.text = "Implementation Schedule (6-Month Plan)"
    header_p = header_tf.paragraphs[0]
    header_p.alignment = PP_ALIGN.LEFT
    run = header_p.runs[0] if header_p.runs else header_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = HEADING_SIZE
    run.font.bold = True
    run.font.color.rgb = ColorPalette.HEADING_TEXT

    # スケジュールコンテンツ
    schedule_box = slide.shapes.add_textbox(
        Inches(1), Inches(1.5),
        Inches(11), Inches(5.5)
    )
    
    schedule_text = [
        "【Phase 1: Requirements & Design (Apr-May)】",
        "• Detailed business requirement analysis.",
        "• System design and architecture finalization.",
        "• Data migration planning.",
        "",
        "【Phase 2: Development & Build (May-Jul)】",
        "• Platform setup and core feature development.",
        "• Implementation of external system integrations.",
        "• User interface development.",
        "",
        "【Phase 3: Testing & Migration (Jul-Aug)】",
        "• Unit and integration testing.",
        "• User acceptance testing (UAT).",
        "• Data migration and system switchover preparation.",
        "",
        "【Phase 4: Go-Live & Stabilization (Sep)】",
        "• Phased production rollout.",
        "• User training sessions.",
        "• Establishment of operational support."
    ]
    apply_body_style(schedule_box, schedule_text, font_size=BODY_SIZE, para_spacing=Pt(12))

    add_footer(slide, prs, include_page_number=True, page_num="5/10")

def create_team_structure(prs):
    """実施体制のスライド (シンプル)"""
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    
    add_background(slide, prs)
    
    # ヘッダーバー
    add_shape(
        slide, MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0),
        prs.slide_width, Inches(1),
        fill_color=ColorPalette.HEADING_BG
    )
    
    # ヘッダータイトル
    header_title = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.2),
        Inches(12), Inches(0.6)
    )
    header_tf = header_title.text_frame
    header_tf.text = "Project Team Structure"
    header_p = header_tf.paragraphs[0]
    header_p.alignment = PP_ALIGN.LEFT
    run = header_p.runs[0] if header_p.runs else header_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = HEADING_SIZE
    run.font.bold = True
    run.font.color.rgb = ColorPalette.HEADING_TEXT

    # 左カラム：推進体制
    team_box = slide.shapes.add_textbox(
        Inches(1), Inches(1.5),
        Inches(5.5), Inches(5)
    )
    team_tf = team_box.text_frame
    team_tf.text = "Core Project Team"
    team_p = team_tf.paragraphs[0]
    team_p.space_after = Pt(12)
    run = team_p.runs[0] if team_p.runs else team_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = SUBHEADING_SIZE
    run.font.bold = True

    team_text = [
        "• Project Sponsor: Head of Corporate Planning",
        "• Project Manager: IT Department Manager",
        "• Technical Lead: Lead Systems Developer",
        "• Business Process Owners: Representatives from each dept."
    ]
    apply_body_style(team_box, team_text, para_spacing=Pt(12))
    # サブヘッダー部分を空文字に
    team_tf.paragraphs[0].text = ""

    # 右カラム：役割とコミュニケーション
    role_comm_box = slide.shapes.add_textbox(
        Inches(7), Inches(1.5),
        Inches(5.5), Inches(5)
    )
    role_comm_tf = role_comm_box.text_frame
    role_comm_tf.clear()  # テキストフレームを初期化

    # 役割 ヘッダー追加
    role_p = role_comm_tf.add_paragraph()
    role_p.text = "Roles & Responsibilities"
    role_p.space_after = Pt(12)
    run = role_p.runs[0] if role_p.runs else role_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = SUBHEADING_SIZE
    run.font.bold = True

    role_text = [
        "• Req. & Design: Our Consultants + Client Business Experts",
        "• Development: Our Engineering Team (5 members)",
        "• Testing & QA: Our QA Team + Client Testers",
        "• Implementation & Training: Our Support Team"
    ]
    for text in role_text:
        p = role_comm_tf.add_paragraph()
        if text.startswith("• "):
            p.text = text[2:]
            p.level = 1
        else:
            p.text = text
            p.level = 0
        p.space_after = Pt(12)
        run = p.runs[0] if p.runs else p.add_run()
        run.font.name = BODY_FONT
        run.font.size = BODY_SIZE
        run.font.color.rgb = ColorPalette.TEXT

    # コミュニケーション ヘッダー追加
    comm_p = role_comm_tf.add_paragraph()
    comm_p.text = "Communication Plan"
    comm_p.space_before = Pt(18)
    comm_p.space_after = Pt(12)
    run = comm_p.runs[0] if comm_p.runs else comm_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = SUBHEADING_SIZE
    run.font.bold = True

    comm_text = [
        "• Weekly Progress Meetings (Online)",
        "• Monthly Steering Committee (In-person)",
        "• Daily Stand-ups (Development Team)",
        "• Shared Issue Tracking System"
    ]
    for text in comm_text:
        p = role_comm_tf.add_paragraph()
        if text.startswith("• "):
            p.text = text[2:]
            p.level = 1
        else:
            p.text = text
            p.level = 0
        p.space_after = Pt(12)
        run = p.runs[0] if p.runs else p.add_run()
        run.font.name = BODY_FONT
        run.font.size = BODY_SIZE
        run.font.color.rgb = ColorPalette.TEXT

    add_footer(slide, prs, include_page_number=True, page_num="6/10")

def create_risk_management(prs):
    """リスク管理計画のスライド (シンプル)"""
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    
    add_background(slide, prs)
    
    # ヘッダーバー
    add_shape(
        slide, MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0),
        prs.slide_width, Inches(1),
        fill_color=ColorPalette.HEADING_BG
    )
    
    # ヘッダータイトル
    header_title = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.2),
        Inches(12), Inches(0.6)
    )
    header_tf = header_title.text_frame
    header_tf.text = "Risk Management Plan"
    header_p = header_tf.paragraphs[0]
    header_p.alignment = PP_ALIGN.LEFT
    run = header_p.runs[0] if header_p.runs else header_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = HEADING_SIZE
    run.font.bold = True
    run.font.color.rgb = ColorPalette.HEADING_TEXT

    # リスクコンテンツ
    risk_box = slide.shapes.add_textbox(
        Inches(1), Inches(1.5),
        Inches(11), Inches(5.5)
    )
    
    risk_points = [
        "【Key Risks & Mitigation Strategies】",
        "",
        "1. Risk: Scope Creep / Changes Leading to Delays",
        "   Mitigation: Agile methodology, regular requirement reviews, strict change control.",
        "",
        "2. Risk: Data Loss / Inconsistency During Migration",
        "   Mitigation: Pre-migration data cleansing, phased approach, dual validation.",
        "",
        "3. Risk: Low User Adoption",
        "   Mitigation: Early user involvement, comprehensive training, continuous feedback loop.",
        "",
        "4. Risk: Integration Issues with Existing Systems",
        "   Mitigation: Detailed interface design, phased integration testing, fallback mechanisms.",
        "",
        "5. Risk: Security Incidents",
        "   Mitigation: Security design reviews, vulnerability assessments, incident response plan."
    ]
    apply_body_style(risk_box, risk_points, font_size=BODY_SIZE, para_spacing=Pt(10))
    # 調整：Mitigation の行はレベルを下げ・サイズを少し小さく
    tf = risk_box.text_frame
    for i, p in enumerate(tf.paragraphs):
        if i > 1 and i % 3 == 0:
            p.level = 1
            p.space_before = Pt(0)
            p.space_after = Pt(10)
            if p.runs:
                p.runs[0].font.size = Pt(14)

    add_footer(slide, prs, include_page_number=True, page_num="7/10")

def create_budget(prs):
    """予算計画のスライド (シンプル)"""
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    
    add_background(slide, prs)
    
    # ヘッダーバー
    add_shape(
        slide, MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0),
        prs.slide_width, Inches(1),
        fill_color=ColorPalette.HEADING_BG
    )
    
    # ヘッダータイトル
    header_title = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.2),
        Inches(12), Inches(0.6)
    )
    header_tf = header_title.text_frame
    header_tf.text = "Budget Plan & ROI"
    header_p = header_tf.paragraphs[0]
    header_p.alignment = PP_ALIGN.LEFT
    run = header_p.runs[0] if header_p.runs else header_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = HEADING_SIZE
    run.font.bold = True
    run.font.color.rgb = ColorPalette.HEADING_TEXT

    # 左カラム：コスト
    cost_box = slide.shapes.add_textbox(
        Inches(1), Inches(1.5),
        Inches(5.5), Inches(5)
    )
    cost_tf = cost_box.text_frame
    cost_tf.clear()
    
    cost_p = cost_tf.add_paragraph()
    cost_p.text = "Initial Investment"
    cost_p.space_after = Pt(12)
    run = cost_p.runs[0] if cost_p.runs else cost_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = SUBHEADING_SIZE
    run.font.bold = True

    initial_text = [
        "• Design & Development: ¥20M",
        "• Hardware & Cloud Setup: ¥5M",
        "• Data Migration & Testing: ¥6M",
        "• Training & Support: ¥4M",
        "• Total Initial Cost: ¥35M"
    ]
    for text in initial_text:
        p = cost_tf.add_paragraph()
        if text.startswith("• "):
            p.text = text[2:]
            p.level = 1
        else:
            p.text = text
            p.level = 0
        p.space_after = Pt(10)
        run = p.runs[0] if p.runs else p.add_run()
        run.font.name = BODY_FONT
        run.font.size = BODY_SIZE
        run.font.color.rgb = ColorPalette.TEXT

    run_p = cost_tf.add_paragraph()
    run_p.text = "Annual Running Costs"
    run_p.space_before = Pt(18)
    run_p.space_after = Pt(12)
    run = run_p.runs[0] if run_p.runs else run_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = SUBHEADING_SIZE
    run.font.bold = True

    running_text = [
        "• Cloud Infrastructure: ¥3M",
        "• Licensing Fees: ¥2M",
        "• Maintenance & Support: ¥3M",
        "• Total Annual Cost: ¥8M"
    ]
    for text in running_text:
        p = cost_tf.add_paragraph()
        if text.startswith("• "):
            p.text = text[2:]
            p.level = 1
        else:
            p.text = text
            p.level = 0
        p.space_after = Pt(10)
        run = p.runs[0] if p.runs else p.add_run()
        run.font.name = BODY_FONT
        run.font.size = BODY_SIZE
        run.font.color.rgb = ColorPalette.TEXT

    # 右カラム：ROI
    roi_box = slide.shapes.add_textbox(
        Inches(7), Inches(1.5),
        Inches(5.5), Inches(5)
    )
    roi_tf = roi_box.text_frame
    roi_tf.clear()

    roi_p = roi_tf.add_paragraph()
    roi_p.text = "Return on Investment (ROI)"
    roi_p.space_after = Pt(12)
    run = roi_p.runs[0] if roi_p.runs else roi_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = SUBHEADING_SIZE
    run.font.bold = True

    roi_text = [
        "【Cost Savings】",
        "• Labor cost reduction (efficiency): ¥12M/year",
        "• System consolidation savings: ¥8M/year",
        "",
        "【Qualitative Benefits】",
        "• Faster decision-making",
        "• Improved customer satisfaction",
        "• Strategic advantage through data utilization",
        "",
        "【Payback Period】",
        "• Approx. 18 months"
    ]
    for text in roi_text:
        p = roi_tf.add_paragraph()
        if text.startswith("【") and text.endswith("】"):
            p.text = text
            p.level = 0
            p.space_before = Pt(15)
            run = p.runs[0] if p.runs else p.add_run()
            run.font.bold = True
            run.font.size = BODY_SIZE + Pt(2)
        elif text.startswith("• "):
            p.text = text[2:]
            p.level = 1
            run = p.runs[0] if p.runs else p.add_run()
            run.font.bold = False
        elif text.strip():
            p.text = text
            p.level = 0
            run = p.runs[0] if p.runs else p.add_run()
            run.font.bold = False
        else:
            p.text = ""
        p.space_after = Pt(10)
        if p.runs:
            run = p.runs[0]
            run.font.name = BODY_FONT
            run.font.size = BODY_SIZE
            run.font.color.rgb = ColorPalette.TEXT

    add_footer(slide, prs, include_page_number=True, page_num="8/10")

def create_success_criteria(prs):
    """成功基準と評価方法のスライド (シンプル)"""
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    
    add_background(slide, prs)
    
    # ヘッダーバー
    add_shape(
        slide, MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0),
        prs.slide_width, Inches(1),
        fill_color=ColorPalette.HEADING_BG
    )
    
    # ヘッダータイトル
    header_title = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.2),
        Inches(12), Inches(0.6)
    )
    header_tf = header_title.text_frame
    header_tf.text = "Success Criteria & Evaluation"
    header_p = header_tf.paragraphs[0]
    header_p.alignment = PP_ALIGN.LEFT
    run = header_p.runs[0] if header_p.runs else header_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = HEADING_SIZE
    run.font.bold = True
    run.font.color.rgb = ColorPalette.HEADING_TEXT

    # 成功基準コンテンツ
    criteria_box = slide.shapes.add_textbox(
        Inches(1), Inches(1.5),
        Inches(11), Inches(5.5)
    )
    
    criteria_points = [
        "【Key Performance Indicators (KPIs)】",
        "",
        "System Performance Metrics:",
        "  • Response Time: < 2 seconds (peak)",
        "  • Availability: > 99.9%",
        "  • Concurrent Users: Support up to 300",
        "  • Backup Recovery Time: < 4 hours",
        "",
        "Business Impact Metrics:",
        "  • Process Time Reduction: 30%",
        "  • Customer Response Time Improvement: 50%",
        "  • Data Entry Error Reduction: 90%",
        "  • Paper Usage Reduction: 80%",
        "  • User Satisfaction: > 80%",
        "",
        "【Evaluation Method】",
        "• Quarterly performance measurement reports.",
        "• Monthly user satisfaction surveys.",
        "• Regular tracking of business efficiency metrics.",
        "• Continuous monitoring via real-time dashboards."
    ]
    apply_body_style(criteria_box, criteria_points, font_size=BODY_SIZE, para_spacing=Pt(10))
    tf = criteria_box.text_frame
    for i, p in enumerate(tf.paragraphs):
        if p.text.endswith(":") and not p.text.startswith("【"):
            if p.runs:
                p.runs[0].font.bold = True
                p.space_after = Pt(6)
        if i in [3,4,5,6, 9,10,11,12,13]:
            p.level = 1
            p.space_before = Pt(0)
            p.space_after = Pt(6)
            if p.runs:
                p.runs[0].font.size = Pt(14)

    add_footer(slide, prs, include_page_number=True, page_num="9/10")

def create_conclusion(prs):
    """まとめと次のステップのスライド (ミニマル)"""
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    
    # 背景設定 (黒背景)
    add_background(slide, prs, color=ColorPalette.HEADING_BG)
    
    # タイトル (白文字)
    title_box = slide.shapes.add_textbox(
        Inches(1), Inches(1.5),
        Inches(11), Inches(1)
    )
    title_tf = title_box.text_frame
    title_tf.text = "Conclusion & Next Steps"
    title_p = title_tf.paragraphs[0]
    title_p.alignment = PP_ALIGN.LEFT
    run = title_p.runs[0] if title_p.runs else title_p.add_run()
    run.font.name = TITLE_FONT
    run.font.size = HEADING_SIZE
    run.font.bold = True
    run.font.color.rgb = ColorPalette.HEADING_TEXT

    # まとめコンテンツ (白文字)
    summary_box = slide.shapes.add_textbox(
        Inches(1), Inches(2.8),
        Inches(11), Inches(1.5)
    )
    summary_text = [
        "【Proposal Summary】",
        "• Implement cloud-based system for 30% efficiency gain.",
        "• Phased 6-month rollout minimizes business disruption.",
        "• Investment: ¥35M initial, ¥8M annual. ROI within 18 months."
    ]
    apply_body_style(summary_box, summary_text, color=ColorPalette.HEADING_TEXT, para_spacing=Pt(12))

    # 次のステップコンテンツ (白文字)
    next_steps_box = slide.shapes.add_textbox(
        Inches(1), Inches(4.5),
        Inches(11), Inches(1.5)
    )
    next_text = [
        "【Next Steps】",
        "• Final review and approval of proposal (within 1 week).",
        "• Project kick-off meeting (within 2 weeks of approval).",
        "• Commence detailed requirements definition (First week of April)."
    ]
    apply_body_style(next_steps_box, next_text, color=ColorPalette.HEADING_TEXT, para_spacing=Pt(12))

    # 連絡先 (白文字)
    contact_box = slide.shapes.add_textbox(
        Inches(1), Inches(6.5),
        Inches(11), Inches(0.5)
    )
    contact_tf = contact_box.text_frame
    contact_tf.text = "Contact: Taro Yamada | yamada.taro@example.com | 03-1234-5678"
    contact_p = contact_tf.paragraphs[0]
    contact_p.alignment = PP_ALIGN.LEFT
    run = contact_p.runs[0] if contact_p.runs else contact_p.add_run()
    run.font.name = BODY_FONT
    run.font.size = BODY_SIZE
    run.font.color.rgb = ColorPalette.FOOTER_BG  # やや薄い白

def main():
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == 'ppt':
        ppt()
    else:
        print("使用方法: doer ppt")
    print("Doerは仕事を完了しました。")

def ppt():
    print("doer ppt が実行されました")

if __name__ == '__main__':
    main()
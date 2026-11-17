"""PowerPoint and MSO COM enums + lookup tables.

Lifted from ppt-mcp (ykuwai/ppt-mcp src/ppt_com/constants.py). Only the
subsets required by the active stage of the controller are mirrored here;
additional groups will be appended as new domains come online.
"""

# ─────────────────────────────────────────────────────────────
# MsoTriState
# ─────────────────────────────────────────────────────────────
msoTrue = -1
msoFalse = 0
msoCTrue = 1
msoTriStateToggle = -3
msoTriStateMixed = -2

# ─────────────────────────────────────────────────────────────
# PpWindowState
# ─────────────────────────────────────────────────────────────
ppWindowNormal = 1
ppWindowMinimized = 2
ppWindowMaximized = 3

WINDOW_STATE_NAMES = {
    ppWindowNormal: "normal",
    ppWindowMinimized: "minimized",
    ppWindowMaximized: "maximized",
}
WINDOW_STATE_VALUES = {v: k for k, v in WINDOW_STATE_NAMES.items()}

# ─────────────────────────────────────────────────────────────
# PpViewType
# ─────────────────────────────────────────────────────────────
ppViewSlide = 1
ppViewSlideMaster = 2
ppViewNotesPage = 3
ppViewNormal = 9
ppViewSlideSorter = 7

# ─────────────────────────────────────────────────────────────
# PpSlideLayout
# ─────────────────────────────────────────────────────────────
ppLayoutTitle = 1
ppLayoutText = 2
ppLayoutTwoColumnText = 3
ppLayoutTable = 4
ppLayoutTextAndChart = 5
ppLayoutChartAndText = 6
ppLayoutChart = 8
ppLayoutTitleOnly = 11
ppLayoutBlank = 12
ppLayoutObject = 16
ppLayoutLargeObject = 15
ppLayoutCustom = 32
ppLayoutSectionHeader = 33
ppLayoutComparison = 34
ppLayoutContentWithCaption = 35
ppLayoutPictureWithCaption = 36

# Friendly layout name → PpSlideLayout int
LAYOUT_NAME_MAP = {
    "title": ppLayoutTitle,
    "text": ppLayoutText,
    "two_column_text": ppLayoutTwoColumnText,
    "table": ppLayoutTable,
    "chart": ppLayoutChart,
    "title_only": ppLayoutTitleOnly,
    "blank": ppLayoutBlank,
    "object": ppLayoutObject,
    "large_object": ppLayoutLargeObject,
    "section_header": ppLayoutSectionHeader,
    "comparison": ppLayoutComparison,
    "content_with_caption": ppLayoutContentWithCaption,
    "picture_with_caption": ppLayoutPictureWithCaption,
}

# ─────────────────────────────────────────────────────────────
# PpSaveAsFileType
# ─────────────────────────────────────────────────────────────
ppSaveAsPresentation = 1
ppSaveAsTemplate = 5
ppSaveAsShow = 7
ppSaveAsDefault = 11
ppSaveAsGIF = 16
ppSaveAsJPG = 17
ppSaveAsPNG = 18
ppSaveAsBMP = 19
ppSaveAsTIF = 21
ppSaveAsOpenXMLPresentation = 24
ppSaveAsOpenXMLPresentationMacroEnabled = 25
ppSaveAsOpenXMLTemplate = 26
ppSaveAsOpenXMLShow = 28
ppSaveAsPDF = 32
ppSaveAsXPS = 33

SAVE_FORMAT_MAP = {
    "pptx": ppSaveAsOpenXMLPresentation,
    "pptm": ppSaveAsOpenXMLPresentationMacroEnabled,
    "ppsx": ppSaveAsOpenXMLShow,
    "potx": ppSaveAsOpenXMLTemplate,
    "pdf": ppSaveAsPDF,
    "xps": ppSaveAsXPS,
    "png": ppSaveAsPNG,
    "jpg": ppSaveAsJPG,
    "gif": ppSaveAsGIF,
    "bmp": ppSaveAsBMP,
    "tif": ppSaveAsTIF,
    "default": ppSaveAsDefault,
}

# ─────────────────────────────────────────────────────────────
# Slide size presets (points; 1 inch = 72 points)
# ─────────────────────────────────────────────────────────────
SLIDE_WIDTH_16_9 = 960.0
SLIDE_HEIGHT_16_9 = 540.0
SLIDE_WIDTH_4_3 = 720.0
SLIDE_HEIGHT_4_3 = 540.0

SLIDE_SIZE_PRESETS = {
    "16:9": (SLIDE_WIDTH_16_9, SLIDE_HEIGHT_16_9),
    "4:3": (SLIDE_WIDTH_4_3, SLIDE_HEIGHT_4_3),
}

# ─────────────────────────────────────────────────────────────
# PpPlaceholderType (canonical values from MS PpPlaceholderType enum)
# ─────────────────────────────────────────────────────────────
ppPlaceholderMixed = -2
ppPlaceholderTitle = 1
ppPlaceholderBody = 2
ppPlaceholderCenterTitle = 3
ppPlaceholderSubtitle = 4
ppPlaceholderVerticalTitle = 5
ppPlaceholderVerticalBody = 6
ppPlaceholderObject = 7
ppPlaceholderChart = 8
ppPlaceholderBitmap = 9
ppPlaceholderMediaClip = 10
ppPlaceholderOrgChart = 11
ppPlaceholderTable = 12
ppPlaceholderSlideNumber = 13
ppPlaceholderHeader = 14
ppPlaceholderFooter = 15
ppPlaceholderDate = 16
ppPlaceholderVerticalObject = 17
ppPlaceholderPicture = 18

# Friendly name → tuple of acceptable PpPlaceholderType values.
# Multiple values matter: PowerPoint's "Title Slide" layout uses
# ppPlaceholderCenterTitle (3), while "Title and Content" uses
# ppPlaceholderTitle (1). Users semantically mean "the title text" in
# both cases, so `title` matches either.
PLACEHOLDER_NAME_MAP: dict[str, tuple[int, ...]] = {
    "title": (ppPlaceholderTitle, ppPlaceholderCenterTitle, ppPlaceholderVerticalTitle),
    "center_title": (ppPlaceholderCenterTitle,),
    "body": (ppPlaceholderBody, ppPlaceholderVerticalBody),
    "subtitle": (ppPlaceholderSubtitle,),
    "chart": (ppPlaceholderChart,),
    "table": (ppPlaceholderTable,),
    "picture": (ppPlaceholderPicture, ppPlaceholderBitmap),
    "media": (ppPlaceholderMediaClip,),
    "object": (ppPlaceholderObject, ppPlaceholderVerticalObject),
    "header": (ppPlaceholderHeader,),
    "footer": (ppPlaceholderFooter,),
    "date": (ppPlaceholderDate,),
    "slide_number": (ppPlaceholderSlideNumber,),
    "org_chart": (ppPlaceholderOrgChart,),
}

PLACEHOLDER_TYPE_NAMES = {
    ppPlaceholderTitle: "title",
    ppPlaceholderBody: "body",
    ppPlaceholderCenterTitle: "center_title",
    ppPlaceholderSubtitle: "subtitle",
    ppPlaceholderVerticalTitle: "vertical_title",
    ppPlaceholderVerticalBody: "vertical_body",
    ppPlaceholderObject: "object",
    ppPlaceholderChart: "chart",
    ppPlaceholderBitmap: "bitmap",
    ppPlaceholderMediaClip: "media",
    ppPlaceholderOrgChart: "org_chart",
    ppPlaceholderTable: "table",
    ppPlaceholderSlideNumber: "slide_number",
    ppPlaceholderHeader: "header",
    ppPlaceholderFooter: "footer",
    ppPlaceholderDate: "date",
    ppPlaceholderVerticalObject: "vertical_object",
    ppPlaceholderPicture: "picture",
    ppPlaceholderMixed: "mixed",
}

# ─────────────────────────────────────────────────────────────
# MsoShapeType (partial — most common shapes for snapshot listing)
# ─────────────────────────────────────────────────────────────
msoAutoShape = 1
msoCallout = 2
msoChart = 3
msoComment = 4
msoFreeform = 5
msoGroup = 6
msoEmbeddedOLEObject = 7
msoLine = 9
msoLinkedOLEObject = 10
msoLinkedPicture = 11
msoOLEControlObject = 12
msoPicture = 13
msoPlaceholder = 14
msoTextBox = 17
msoMedia = 16
msoTable = 19
msoCanvas = 20
msoDiagram = 21
msoInk = 22
msoInkComment = 23
msoSmartArt = 24
msoIgxGraphic = 24  # alias

SHAPE_TYPE_NAMES = {
    msoAutoShape: "auto_shape",
    msoCallout: "callout",
    msoChart: "chart",
    msoComment: "comment",
    msoFreeform: "freeform",
    msoGroup: "group",
    msoEmbeddedOLEObject: "embedded_ole",
    msoLine: "line",
    msoLinkedOLEObject: "linked_ole",
    msoLinkedPicture: "linked_picture",
    msoOLEControlObject: "ole_control",
    msoPicture: "picture",
    msoPlaceholder: "placeholder",
    msoTextBox: "text_box",
    msoMedia: "media",
    msoTable: "table",
    msoCanvas: "canvas",
    msoDiagram: "diagram",
    msoInk: "ink",
    msoInkComment: "ink_comment",
    msoSmartArt: "smart_art",
}

# ─────────────────────────────────────────────────────────────
# MsoTextOrientation (for AddTextbox)
# ─────────────────────────────────────────────────────────────
msoTextOrientationHorizontal = 1
msoTextOrientationUpward = 2
msoTextOrientationDownward = 3
msoTextOrientationVerticalFarEast = 4
msoTextOrientationVertical = 5

TEXT_ORIENTATION_MAP = {
    "horizontal": msoTextOrientationHorizontal,
    "vertical": msoTextOrientationVertical,
    "upward": msoTextOrientationUpward,
    "downward": msoTextOrientationDownward,
}

# ─────────────────────────────────────────────────────────────
# MsoAutoShapeType (curated subset — the shapes agents actually
# request. ~80 of ~180 total. Add more on demand.)
# ─────────────────────────────────────────────────────────────
msoShapeRectangle = 1
msoShapeParallelogram = 2
msoShapeTrapezoid = 3
msoShapeDiamond = 4
msoShapeRoundedRectangle = 5
msoShapeOctagon = 6
msoShapeIsoscelesTriangle = 7
msoShapeRightTriangle = 8
msoShapeOval = 9
msoShapeHexagon = 10
msoShapeCross = 11
msoShapeRegularPentagon = 12
msoShapeCan = 13
msoShapeCube = 14
msoShapeBevel = 15
msoShapeFoldedCorner = 16
msoShapeSmileyFace = 17
msoShapeDonut = 18
msoShapeNoSymbol = 19
msoShapeBlockArc = 20
msoShapeHeart = 21
msoShapeLightningBolt = 22
msoShapeSun = 23
msoShapeMoon = 24
msoShapeArc = 25
msoShapeDoubleBracket = 26
msoShapeDoubleBrace = 27
msoShapePlaque = 28
msoShapeLeftBracket = 29
msoShapeRightBracket = 30
msoShapeLeftBrace = 31
msoShapeRightBrace = 32
msoShapeRightArrow = 33
msoShapeLeftArrow = 34
msoShapeUpArrow = 35
msoShapeDownArrow = 36
msoShapeLeftRightArrow = 37
msoShapeUpDownArrow = 38
msoShapeQuadArrow = 39
msoShapeLeftRightUpArrow = 40
msoShapeBentArrow = 41
msoShapeUTurnArrow = 42
msoShapeLeftUpArrow = 43
msoShapeBentUpArrow = 44
msoShapeCurvedRightArrow = 45
msoShapeCurvedLeftArrow = 46
msoShapeCurvedUpArrow = 47
msoShapeCurvedDownArrow = 48
msoShapeStripedRightArrow = 49
msoShapeNotchedRightArrow = 50
msoShapePentagon = 51
msoShapeChevron = 52
msoShapeRightArrowCallout = 53
msoShapeLeftArrowCallout = 54
msoShapeUpArrowCallout = 55
msoShapeDownArrowCallout = 56
msoShapeLeftRightArrowCallout = 57
msoShapeUpDownArrowCallout = 58
msoShapeQuadArrowCallout = 59
msoShapeFlowchartProcess = 61
msoShapeFlowchartAlternateProcess = 62
msoShapeFlowchartDecision = 63
msoShapeFlowchartData = 64
msoShapeFlowchartPredefinedProcess = 65
msoShapeFlowchartInternalStorage = 66
msoShapeFlowchartDocument = 67
msoShapeFlowchartMultidocument = 68
msoShapeFlowchartTerminator = 69
msoShapeFlowchartPreparation = 70
msoShapeFlowchartManualInput = 71
msoShapeFlowchartManualOperation = 72
msoShapeFlowchartConnector = 73
msoShapeFlowchartOffpageConnector = 74
msoShape5pointStar = 92
msoShape8pointStar = 94
msoShape16pointStar = 95
msoShape24pointStar = 96
msoShape32pointStar = 97
msoShapeRectangularCallout = 105
msoShapeRoundedRectangularCallout = 106
msoShapeOvalCallout = 107
msoShapeCloudCallout = 108
msoShapeLineCallout1 = 109
msoShapeLineCallout2 = 110

AUTOSHAPE_NAME_MAP: dict[str, int] = {
    "rectangle": msoShapeRectangle,
    "rounded_rectangle": msoShapeRoundedRectangle,
    "oval": msoShapeOval,
    "ellipse": msoShapeOval,
    "circle": msoShapeOval,
    "diamond": msoShapeDiamond,
    "triangle": msoShapeIsoscelesTriangle,
    "isoceles_triangle": msoShapeIsoscelesTriangle,
    "right_triangle": msoShapeRightTriangle,
    "parallelogram": msoShapeParallelogram,
    "trapezoid": msoShapeTrapezoid,
    "pentagon": msoShapeRegularPentagon,
    "hexagon": msoShapeHexagon,
    "octagon": msoShapeOctagon,
    "cross": msoShapeCross,
    "plus": msoShapeCross,
    "can": msoShapeCan,
    "cube": msoShapeCube,
    "bevel": msoShapeBevel,
    "smiley": msoShapeSmileyFace,
    "smiley_face": msoShapeSmileyFace,
    "heart": msoShapeHeart,
    "lightning": msoShapeLightningBolt,
    "lightning_bolt": msoShapeLightningBolt,
    "sun": msoShapeSun,
    "moon": msoShapeMoon,
    "arc": msoShapeArc,
    "donut": msoShapeDonut,
    "no_symbol": msoShapeNoSymbol,
    "plaque": msoShapePlaque,
    "right_arrow": msoShapeRightArrow,
    "left_arrow": msoShapeLeftArrow,
    "up_arrow": msoShapeUpArrow,
    "down_arrow": msoShapeDownArrow,
    "left_right_arrow": msoShapeLeftRightArrow,
    "up_down_arrow": msoShapeUpDownArrow,
    "quad_arrow": msoShapeQuadArrow,
    "bent_arrow": msoShapeBentArrow,
    "u_turn_arrow": msoShapeUTurnArrow,
    "curved_right_arrow": msoShapeCurvedRightArrow,
    "curved_left_arrow": msoShapeCurvedLeftArrow,
    "curved_up_arrow": msoShapeCurvedUpArrow,
    "curved_down_arrow": msoShapeCurvedDownArrow,
    "striped_right_arrow": msoShapeStripedRightArrow,
    "notched_right_arrow": msoShapeNotchedRightArrow,
    "chevron": msoShapeChevron,
    "star": msoShape5pointStar,
    "star_5": msoShape5pointStar,
    "star_8": msoShape8pointStar,
    "star_16": msoShape16pointStar,
    "star_24": msoShape24pointStar,
    "star_32": msoShape32pointStar,
    "rect_callout": msoShapeRectangularCallout,
    "rectangular_callout": msoShapeRectangularCallout,
    "rounded_callout": msoShapeRoundedRectangularCallout,
    "oval_callout": msoShapeOvalCallout,
    "cloud_callout": msoShapeCloudCallout,
    "line_callout": msoShapeLineCallout1,
    "flowchart_process": msoShapeFlowchartProcess,
    "flowchart_decision": msoShapeFlowchartDecision,
    "flowchart_data": msoShapeFlowchartData,
    "flowchart_document": msoShapeFlowchartDocument,
    "flowchart_terminator": msoShapeFlowchartTerminator,
    "flowchart_connector": msoShapeFlowchartConnector,
}

# ─────────────────────────────────────────────────────────────
# PpParagraphAlignment
# ─────────────────────────────────────────────────────────────
ppAlignLeft = 1
ppAlignCenter = 2
ppAlignRight = 3
ppAlignJustify = 4
ppAlignDistribute = 5

PARAGRAPH_ALIGN_MAP: dict[str, int] = {
    "left": ppAlignLeft,
    "center": ppAlignCenter,
    "right": ppAlignRight,
    "justify": ppAlignJustify,
    "distribute": ppAlignDistribute,
}

# ─────────────────────────────────────────────────────────────
# MsoBulletType (used to disable bullets — disabling = msoBulletNone)
# ─────────────────────────────────────────────────────────────
ppBulletNone = 0
ppBulletUnnumbered = 1
ppBulletNumbered = 2
ppBulletPicture = 3

BULLET_TYPE_MAP: dict[str, int] = {
    "none": ppBulletNone,
    "off": ppBulletNone,
    "unnumbered": ppBulletUnnumbered,
    "bullet": ppBulletUnnumbered,
    "numbered": ppBulletNumbered,
    "number": ppBulletNumbered,
}

# ─────────────────────────────────────────────────────────────
# MsoZOrderCmd
# ─────────────────────────────────────────────────────────────
msoBringToFront = 0
msoSendToBack = 1
msoBringForward = 2
msoSendBackward = 3
msoBringInFrontOfText = 4
msoSendBehindText = 5

Z_ORDER_MAP: dict[str, int] = {
    "bring_to_front": msoBringToFront,
    "send_to_back": msoSendToBack,
    "bring_forward": msoBringForward,
    "send_backward": msoSendBackward,
    "front": msoBringToFront,
    "back": msoSendToBack,
    "forward": msoBringForward,
    "backward": msoSendBackward,
}

# ─────────────────────────────────────────────────────────────
# MsoAlignCmd
# ─────────────────────────────────────────────────────────────
msoAlignLefts = 0
msoAlignCenters = 1
msoAlignRights = 2
msoAlignTops = 3
msoAlignMiddles = 4
msoAlignBottoms = 5

ALIGN_MAP: dict[str, int] = {
    "left": msoAlignLefts,
    "lefts": msoAlignLefts,
    "center": msoAlignCenters,
    "centers": msoAlignCenters,
    "right": msoAlignRights,
    "rights": msoAlignRights,
    "top": msoAlignTops,
    "tops": msoAlignTops,
    "middle": msoAlignMiddles,
    "middles": msoAlignMiddles,
    "bottom": msoAlignBottoms,
    "bottoms": msoAlignBottoms,
}

# Distribute (MsoDistributeCmd)
msoDistributeHorizontally = 0
msoDistributeVertically = 1

DISTRIBUTE_MAP: dict[str, int] = {
    "horizontally": msoDistributeHorizontally,
    "horizontal": msoDistributeHorizontally,
    "vertically": msoDistributeVertically,
    "vertical": msoDistributeVertically,
}

# ─────────────────────────────────────────────────────────────
# MsoLineDashStyle
# ─────────────────────────────────────────────────────────────
msoLineSolid = 1
msoLineSquareDot = 2
msoLineRoundDot = 3
msoLineDash = 4
msoLineDashDot = 5
msoLineLongDash = 7
msoLineLongDashDot = 8
msoLineDashDotDot = 9
msoLineLongDashDotDot = 10

DASH_STYLE_MAP: dict[str, int] = {
    "solid": msoLineSolid,
    "square_dot": msoLineSquareDot,
    "round_dot": msoLineRoundDot,
    "dot": msoLineRoundDot,
    "dash": msoLineDash,
    "dash_dot": msoLineDashDot,
    "long_dash": msoLineLongDash,
    "long_dash_dot": msoLineLongDashDot,
    "dash_dot_dot": msoLineDashDotDot,
    "long_dash_dot_dot": msoLineLongDashDotDot,
}

# ─────────────────────────────────────────────────────────────
# MsoConnectorType
# ─────────────────────────────────────────────────────────────
msoConnectorStraight = 1
msoConnectorElbow = 2
msoConnectorCurve = 3

CONNECTOR_TYPE_MAP: dict[str, int] = {
    "straight": msoConnectorStraight,
    "line": msoConnectorStraight,
    "elbow": msoConnectorElbow,
    "right_angle": msoConnectorElbow,
    "curve": msoConnectorCurve,
    "curved": msoConnectorCurve,
}

# ─────────────────────────────────────────────────────────────
# XlChartType (curated subset — most common chart types)
# ─────────────────────────────────────────────────────────────
xlColumnClustered = 51
xlColumnStacked = 52
xlColumnStacked100 = 53
xlBarClustered = 57
xlBarStacked = 58
xlLine = 4
xlLineMarkers = 65
xlPie = 5
xlDoughnut = -4120
xlArea = 1
xlAreaStacked = 76
xlXYScatter = -4169
xlXYScatterLines = 74
xlRadar = -4151
xl3DColumn = -4100
xl3DPie = -4102
xl3DLine = -4101

CHART_TYPE_MAP: dict[str, int] = {
    "column": xlColumnClustered,
    "column_clustered": xlColumnClustered,
    "column_stacked": xlColumnStacked,
    "column_stacked_100": xlColumnStacked100,
    "bar": xlBarClustered,
    "bar_clustered": xlBarClustered,
    "bar_stacked": xlBarStacked,
    "line": xlLine,
    "line_markers": xlLineMarkers,
    "pie": xlPie,
    "doughnut": xlDoughnut,
    "donut": xlDoughnut,
    "area": xlArea,
    "area_stacked": xlAreaStacked,
    "scatter": xlXYScatter,
    "scatter_lines": xlXYScatterLines,
    "radar": xlRadar,
    "column_3d": xl3DColumn,
    "pie_3d": xl3DPie,
    "line_3d": xl3DLine,
}

# ─────────────────────────────────────────────────────────────
# XlAxisType / XlLegendPosition / XlDataLabelsType — chart polish
# ─────────────────────────────────────────────────────────────
xlCategory = 1
xlValue = 2
xlSeriesAxis = 3

AXIS_TYPE_MAP: dict[str, int] = {
    "x": xlCategory,
    "category": xlCategory,
    "horizontal": xlCategory,
    "y": xlValue,
    "value": xlValue,
    "vertical": xlValue,
    "series": xlSeriesAxis,
}

xlLegendPositionBottom = -4107
xlLegendPositionCorner = 2
xlLegendPositionLeft = -4131
xlLegendPositionRight = -4152
xlLegendPositionTop = -4160

LEGEND_POSITION_MAP: dict[str, int] = {
    "bottom": xlLegendPositionBottom,
    "top": xlLegendPositionTop,
    "left": xlLegendPositionLeft,
    "right": xlLegendPositionRight,
    "corner": xlLegendPositionCorner,
}

xlDataLabelsShowNone = -4142
xlDataLabelsShowValue = 2
xlDataLabelsShowPercent = 3
xlDataLabelsShowLabel = 4
xlDataLabelsShowLabelAndPercent = 5

DATA_LABEL_TYPE_MAP: dict[str, int] = {
    "none": xlDataLabelsShowNone,
    "off": xlDataLabelsShowNone,
    "value": xlDataLabelsShowValue,
    "percent": xlDataLabelsShowPercent,
    "label": xlDataLabelsShowLabel,
    "label_percent": xlDataLabelsShowLabelAndPercent,
    "label_and_percent": xlDataLabelsShowLabelAndPercent,
}

# ─────────────────────────────────────────────────────────────
# MsoAnchor — vertical alignment inside text frames / table cells
# ─────────────────────────────────────────────────────────────
msoAnchorTop = 1
msoAnchorMiddle = 3
msoAnchorBottom = 4

VERTICAL_ALIGN_MAP: dict[str, int] = {
    "top": msoAnchorTop,
    "middle": msoAnchorMiddle,
    "center": msoAnchorMiddle,
    "bottom": msoAnchorBottom,
}

# ─────────────────────────────────────────────────────────────
# Table style GUIDs (curated — full set is ~75 styles, surfaced by
# friendly name. Raw GUID strings also accepted by table_apply_style.)
# ─────────────────────────────────────────────────────────────
TABLE_STYLE_MAP: dict[str, str] = {
    "none": "{5940675A-B579-460E-94D1-54222C63F5DA}",
    "no_style_no_grid": "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
    "no_style_grid": "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
    "themed_style_1_accent_1": "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}",
    "themed_style_1_accent_2": "{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}",
    "themed_style_1_accent_3": "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}",
    "themed_style_1_accent_4": "{D27102A9-8310-4765-A935-A1911B00CA55}",
    "themed_style_1_accent_5": "{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}",
    "themed_style_1_accent_6": "{68D230F3-CF80-4859-8CE7-A43EE81993B5}",
    "light_style_1": "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}",
    "light_style_2": "{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}",
    "light_style_2_accent_1": "{69C7853C-536D-4A76-A0AE-DD22124D55A5}",
    "light_style_2_accent_2": "{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}",
    "medium_style_1": "{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}",
    "medium_style_2": "{5DA37D80-6434-44D0-A028-1B22A696006F}",
    "medium_style_2_accent_1": "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}",
    "medium_style_2_accent_2": "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}",
    "medium_style_2_accent_3": "{00A15C55-8517-42AA-B614-E9B94910E393}",
    "medium_style_2_accent_4": "{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}",
    "medium_style_2_accent_5": "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}",
    "medium_style_2_accent_6": "{35758FB7-9AC5-4552-8A53-C91805E547FA}",
    "medium_style_3": "{2A488322-F2BA-4B5B-9748-0D474271808F}",
    "medium_style_4": "{D7AC3CCA-C797-4891-BE02-D94E43425B78}",
    "dark_style_1": "{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}",
    "dark_style_2": "{5202B0CA-FC54-4496-8BCA-5EF66A818D29}",
}

# ─────────────────────────────────────────────────────────────
# Valid file extensions for open()
# ─────────────────────────────────────────────────────────────
VALID_EXTENSIONS = {
    ".pptx",
    ".pptm",
    ".ppt",
    ".ppsx",
    ".pps",
    ".potx",
    ".potm",
    ".pot",
}

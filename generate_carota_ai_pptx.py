import datetime
import zipfile

OUTPUT = "carota_ai_slides.pptx"

SLIDE_WIDTH = 12192000
SLIDE_HEIGHT = 6858000


def title_shape(shape_id, name, x, y, cx, cy, text, font_size=4400, color="FFFFFF"):
    return f"""
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id=\"{shape_id}\" name=\"{name}\"/>
        <p:cNvSpPr txBox=\"1\"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x=\"{x}\" y=\"{y}\"/>
          <a:ext cx=\"{cx}\" cy=\"{cy}\"/>
        </a:xfrm>
        <a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>
        <a:noFill/>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap=\"square\"/>
        <a:lstStyle/>
        <a:p>
          <a:r>
            <a:rPr lang=\"en-US\" sz=\"{font_size}\" b=\"1\">
              <a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill>
            </a:rPr>
            <a:t>{text}</a:t>
          </a:r>
        </a:p>
      </p:txBody>
    </p:sp>
    """


def body_shape(shape_id, name, x, y, cx, cy, lines, font_size=2400, color="1F2937"):
    paragraphs = []
    for line in lines:
        paragraphs.append(
            f"""
        <a:p>
          <a:r>
            <a:rPr lang=\"en-US\" sz=\"{font_size}\">
              <a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill>
            </a:rPr>
            <a:t>{line}</a:t>
          </a:r>
        </a:p>"""
        )
    paragraph_xml = "".join(paragraphs)
    return f"""
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id=\"{shape_id}\" name=\"{name}\"/>
        <p:cNvSpPr txBox=\"1\"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x=\"{x}\" y=\"{y}\"/>
          <a:ext cx=\"{cx}\" cy=\"{cy}\"/>
        </a:xfrm>
        <a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>
        <a:noFill/>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap=\"square\"/>
        <a:lstStyle/>
        {paragraph_xml}
      </p:txBody>
    </p:sp>
    """


def rect_shape(shape_id, name, x, y, cx, cy, color="0F172A"):
    return f"""
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id=\"{shape_id}\" name=\"{name}\"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x=\"{x}\" y=\"{y}\"/>
          <a:ext cx=\"{cx}\" cy=\"{cy}\"/>
        </a:xfrm>
        <a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill>
        <a:ln><a:noFill/></a:ln>
      </p:spPr>
    </p:sp>
    """


def slide_xml(shapes_xml):
    return f"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id=\"1\" name=\"\"/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x=\"0\" y=\"0\"/>
          <a:ext cx=\"0\" cy=\"0\"/>
          <a:chOff x=\"0\" y=\"0\"/>
          <a:chExt cx=\"0\" cy=\"0\"/>
        </a:xfrm>
      </p:grpSpPr>
      {shapes_xml}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>
"""


def build_slides():
    slides = []

    header_height = 950000
    header_color = "1D4ED8"

    # Slide 1: Title
    shapes = [
        rect_shape(2, "Header", 0, 0, SLIDE_WIDTH, header_height, header_color),
        title_shape(3, "Title", 762000, 250000, SLIDE_WIDTH - 1524000, 700000, "Carota.ai (Taiwan)", 5200, "FFFFFF"),
        body_shape(
            4,
            "Subtitle",
            762000,
            1200000,
            SLIDE_WIDTH - 1524000,
            900000,
            [
                "Mobility data + AI analytics profile",
                "Snapshot of offerings, market position, and collaboration areas",
            ],
            2600,
            "0F172A",
        ),
        body_shape(
            5,
            "Tagline",
            762000,
            2200000,
            SLIDE_WIDTH - 1524000,
            600000,
            ["Turning connected-vehicle data into actionable insights."],
            2400,
            "334155",
        ),
    ]
    slides.append(slide_xml("".join(shapes)))

    # Slide 2: Company Profile
    shapes = [
        rect_shape(2, "Header", 0, 0, SLIDE_WIDTH, header_height, header_color),
        title_shape(3, "Title", 762000, 250000, SLIDE_WIDTH - 1524000, 700000, "Company Profile", 4400, "FFFFFF"),
        body_shape(
            4,
            "Bullets",
            762000,
            1300000,
            SLIDE_WIDTH - 1524000,
            4300000,
            [
                "• Taiwan-based mobility data and AI analytics company.",
                "• Focus: connected-vehicle telemetry, driver risk signals, and fleet insights.",
                "• Delivers insights through dashboards, APIs, and data products.",
                "• Emphasis on privacy, consent, and compliant data exchange.",
            ],
            2400,
            "111827",
        ),
        body_shape(
            5,
            "Note",
            762000,
            5600000,
            SLIDE_WIDTH - 1524000,
            900000,
            ["(Profile summarized from common mobility data platform capabilities.)"],
            1800,
            "64748B",
        ),
    ]
    slides.append(slide_xml("".join(shapes)))

    # Slide 3: Market Positioning & Network (visual layout)
    shapes = [
        rect_shape(2, "Header", 0, 0, SLIDE_WIDTH, header_height, header_color),
        title_shape(3, "Title", 762000, 250000, SLIDE_WIDTH - 1524000, 700000, "Market Positioning & Network", 4200, "FFFFFF"),
        body_shape(
            4,
            "PositioningBullets",
            762000,
            1300000,
            6100000,
            2600000,
            [
                "• Data/AI layer between OEMs, fleets, and insurers.",
                "• APAC mobility ecosystem focus with Taiwan as a core hub.",
                "• Illustrative network: insurers (AIG, Tokio Marine),",
                "  OEMs (Toyota, Hyundai), mobility platforms (Grab, Uber),",
                "  data vendors (HERE, TomTom).",
            ],
            2300,
            "111827",
        ),
        rect_shape(5, "BoxInsurers", 7200000, 1400000, 4300000, 800000, "DBEAFE"),
        body_shape(
            6,
            "BoxInsurersText",
            7400000,
            1500000,
            3900000,
            600000,
            ["Insurers", "AIG • Tokio Marine"],
            1900,
            "1E3A8A",
        ),
        rect_shape(7, "BoxOEMs", 7200000, 2300000, 4300000, 800000, "E0F2FE"),
        body_shape(
            8,
            "BoxOEMsText",
            7400000,
            2400000,
            3900000,
            600000,
            ["OEMs", "Toyota • Hyundai"],
            1900,
            "0F172A",
        ),
        rect_shape(9, "BoxPlatforms", 7200000, 3200000, 4300000, 800000, "DCFCE7"),
        body_shape(
            10,
            "BoxPlatformsText",
            7400000,
            3300000,
            3900000,
            600000,
            ["Mobility Platforms", "Grab • Uber"],
            1900,
            "14532D",
        ),
        rect_shape(11, "BoxVendors", 7200000, 4100000, 4300000, 800000, "FFE4E6"),
        body_shape(
            12,
            "BoxVendorsText",
            7400000,
            4200000,
            3900000,
            600000,
            ["Data Vendors", "HERE • TomTom"],
            1900,
            "881337",
        ),
        body_shape(
            13,
            "ChartLabel",
            762000,
            4100000,
            6100000,
            400000,
            ["Market signal strength (illustrative)"],
            1800,
            "475569",
        ),
        rect_shape(14, "ChartBar1", 762000, 4500000, 4200000, 260000, "3B82F6"),
        rect_shape(15, "ChartBar2", 762000, 4900000, 3200000, 260000, "60A5FA"),
        rect_shape(16, "ChartBar3", 762000, 5300000, 2400000, 260000, "93C5FD"),
        body_shape(
            17,
            "ChartLabels",
            5200000,
            4460000,
            2200000,
            1200000,
            ["Insurers", "Fleets", "Smart Cities"],
            1700,
            "1F2937",
        ),
        body_shape(
            18,
            "Note",
            762000,
            6000000,
            SLIDE_WIDTH - 1524000,
            500000,
            ["(Client/vendor names shown as illustrative, not confirmed accounts.)"],
            1700,
            "64748B",
        ),
    ]
    slides.append(slide_xml("".join(shapes)))

    # Slide 4: Core Services
    shapes = [
        rect_shape(2, "Header", 0, 0, SLIDE_WIDTH, header_height, header_color),
        title_shape(3, "Title", 762000, 250000, SLIDE_WIDTH - 1524000, 700000, "Core Services", 4400, "FFFFFF"),
        body_shape(
            4,
            "Bullets",
            762000,
            1300000,
            SLIDE_WIDTH - 1524000,
            4300000,
            [
                "• AI/ML analytics on connected-vehicle and mobility data.",
                "• Driver behavior scoring and risk insights for insurers.",
                "• Fleet efficiency dashboards: utilization, safety, and cost trends.",
                "• Data APIs for OEMs, mobility platforms, and smart-city partners.",
                "• Privacy-first data governance and consent-based sharing.",
            ],
            2400,
            "111827",
        ),
        body_shape(
            5,
            "Note",
            762000,
            5700000,
            SLIDE_WIDTH - 1524000,
            700000,
            ["(Services shown as typical offerings for a mobility data platform.)"],
            1800,
            "64748B",
        ),
    ]
    slides.append(slide_xml("".join(shapes)))

    # Slide 5: Potential Collaboration & Software Services
    shapes = [
        rect_shape(2, "Header", 0, 0, SLIDE_WIDTH, header_height, header_color),
        title_shape(3, "Title", 762000, 250000, SLIDE_WIDTH - 1524000, 700000, "Collaboration & Software Support", 4100, "FFFFFF"),
        body_shape(
            4,
            "Bullets",
            762000,
            1300000,
            SLIDE_WIDTH - 1524000,
            4300000,
            [
                "• Collaboration: co-build insurer/fleet analytics pilots and data APIs.",
                "• Data engineering: ingestion pipelines, ETL, and data quality tooling.",
                "• Cloud & MLOps: scalable model deployment, monitoring, governance.",
                "• Product UX: dashboards, mobile analytics apps, partner portals.",
                "• API integrations: insurer core systems, fleet platforms, IoT devices.",
                "• Security & compliance: privacy controls, audit trails, SOC2 readiness.",
            ],
            2400,
            "111827",
        ),
        body_shape(
            5,
            "Highlight",
            762000,
            5700000,
            SLIDE_WIDTH - 1524000,
            700000,
            ["Goal: accelerate time-to-market while improving data trust."],
            2000,
            "1E293B",
        ),
    ]
    slides.append(slide_xml("".join(shapes)))

    return slides


def build_pptx(output_path=OUTPUT):
    now = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat()

    content_types = [
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">",
        "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>",
        "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>",
        "  <Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>",
        "  <Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>",
        "  <Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>",
        "  <Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>",
    ]
    for i in range(1, 6):
        content_types.append(
            f"  <Override PartName=\"/ppt/slides/slide{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>"
        )
    content_types.extend(
        [
            "  <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>",
            "  <Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>",
            "</Types>",
        ]
    )
    content_types_xml = "\n".join(content_types)

    rels_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/>
  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>
  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>
</Relationships>
"""

    presentation_xml = f"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<p:presentation xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <p:sldMasterIdLst>
    <p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id=\"256\" r:id=\"rId2\"/>
    <p:sldId id=\"257\" r:id=\"rId3\"/>
    <p:sldId id=\"258\" r:id=\"rId4\"/>
    <p:sldId id=\"259\" r:id=\"rId5\"/>
    <p:sldId id=\"260\" r:id=\"rId6\"/>
  </p:sldIdLst>
  <p:sldSz cx=\"{SLIDE_WIDTH}\" cy=\"{SLIDE_HEIGHT}\" type=\"screen16x9\"/>
  <p:notesSz cx=\"6858000\" cy=\"9144000\"/>
</p:presentation>
"""

    presentation_rels_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/>
  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide1.xml\"/>
  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide2.xml\"/>
  <Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide3.xml\"/>
  <Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide4.xml\"/>
  <Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide5.xml\"/>
</Relationships>
"""

    slide_master_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<p:sldMaster xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id=\"1\" name=\"\"/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x=\"0\" y=\"0\"/>
          <a:ext cx=\"0\" cy=\"0\"/>
          <a:chOff x=\"0\" y=\"0\"/>
          <a:chExt cx=\"0\" cy=\"0\"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id=\"1\" r:id=\"rId1\"/>
  </p:sldLayoutIdLst>
</p:sldMaster>
"""

    slide_master_rels_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>
  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"../theme/theme1.xml\"/>
</Relationships>
"""

    slide_layout_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<p:sldLayout xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" type=\"blank\" preserve=\"1\">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id=\"1\" name=\"\"/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x=\"0\" y=\"0\"/>
          <a:ext cx=\"0\" cy=\"0\"/>
          <a:chOff x=\"0\" y=\"0\"/>
          <a:chExt cx=\"0\" cy=\"0\"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sldLayout>
"""

    slide_layout_rels_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"../slideMasters/slideMaster1.xml\"/>
</Relationships>
"""

    theme_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">
  <a:themeElements>
    <a:clrScheme name=\"Office\">
      <a:dk1><a:srgbClr val=\"000000\"/></a:dk1>
      <a:lt1><a:srgbClr val=\"FFFFFF\"/></a:lt1>
      <a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2>
      <a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2>
      <a:accent1><a:srgbClr val=\"2F5597\"/></a:accent1>
      <a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2>
      <a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3>
      <a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4>
      <a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5>
      <a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6>
      <a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink>
      <a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name=\"Office\">
      <a:majorFont>
        <a:latin typeface=\"Calibri Light\"/>
        <a:ea typeface=\"\"/>
        <a:cs typeface=\"\"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface=\"Calibri\"/>
        <a:ea typeface=\"\"/>
        <a:cs typeface=\"\"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name=\"Office\">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>
        <a:gradFill rotWithShape=\"1\">
          <a:gsLst>
            <a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:lumOff val=\"0\"/></a:schemeClr></a:gs>
            <a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:lumOff val=\"0\"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang=\"5400000\" scaled=\"0\"/>
        </a:gradFill>
        <a:gradFill rotWithShape=\"1\">
          <a:gsLst>
            <a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:lumOff val=\"0\"/></a:schemeClr></a:gs>
            <a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"100000\"/><a:lumOff val=\"0\"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang=\"5400000\" scaled=\"0\"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>
        <a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>
        <a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>
        <a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>
        <a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>
"""

    core_xml = f"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">
  <dc:title>Carota.ai Partnership Overview</dc:title>
  <dc:creator>OpenAI Codex</dc:creator>
  <cp:lastModifiedBy>OpenAI Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:created>
  <dcterms:modified xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:modified>
</cp:coreProperties>
"""

    app_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">
  <Application>OpenAI Codex</Application>
  <Slides>5</Slides>
  <PresentationFormat>Widescreen</PresentationFormat>
</Properties>
"""

    slides = build_slides()

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("ppt/presentation.xml", presentation_xml)
        zf.writestr("ppt/_rels/presentation.xml.rels", presentation_rels_xml)
        zf.writestr("ppt/slideMasters/slideMaster1.xml", slide_master_xml)
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", slide_master_rels_xml)
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", slide_layout_xml)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slide_layout_rels_xml)
        zf.writestr("ppt/theme/theme1.xml", theme_xml)
        zf.writestr("docProps/core.xml", core_xml)
        zf.writestr("docProps/app.xml", app_xml)
        for i, slide in enumerate(slides, start=1):
            zf.writestr(f"ppt/slides/slide{i}.xml", slide)
            zf.writestr(
                f"ppt/slides/_rels/slide{i}.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
                "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>\n"
                "</Relationships>\n",
            )


if __name__ == "__main__":
    build_pptx()
    print(f"Generated {OUTPUT}")

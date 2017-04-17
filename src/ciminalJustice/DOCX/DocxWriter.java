package ciminalJustice.DOCX;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintStream;
import java.time.*;
import java.time.format.*;
import static java.time.format.DateTimeFormatter.ISO_LOCAL_DATE;
import static java.time.format.DateTimeFormatter.ISO_LOCAL_DATE_TIME;

/**
 *
 * @author orion116
 */
public class DocxWriter
{
    // add 7z checks
    public static PrintStream input;
    public static void writeMETAFile() throws FileNotFoundException
    {
        PrintStream file = new PrintStream(new FileOutputStream("report/[Content_Types].xml"), false);

        file.printf("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n" +
                    "    <Default Extension=\"png\" ContentType=\"image/png\"/>\n" +
                    "    <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n" +
                    "    <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n" +
                    "    <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\n" +
                    "    <Override PartName=\"/customXml/itemProps1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.customXmlProperties+xml\"/>\n" +
                    "    <Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>\n" +
                    "    <Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>\n" +
                    "    <Override PartName=\"/word/webSettings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml\"/>\n" +
                    "    <Override PartName=\"/word/footnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"/>\n" +
                    "    <Override PartName=\"/word/endnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml\"/>\n" +
                    "    <Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>\n" +
                    "    <Override PartName=\"/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/>\n" +
                    "    <Override PartName=\"/word/header2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>\n" +
                    "    <Override PartName=\"/word/footer2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/>\n" +
                    "    <Override PartName=\"/word/fontTable.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml\"/>\n" +
                    "    <Override PartName=\"/word/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\n" +
                    "    <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\n" +
                    "    <Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>\n" +
                    "</Types>"
                    );
        
		file.close();
    }

    public static void writeHeader1File() throws FileNotFoundException
    {
        PrintStream file = new PrintStream(new FileOutputStream("report/word/header1.xml"), false);

        file.printf("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                    "<w:hdr xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://purl.oclc.org/ooxml/officeDocument/relationships\" xmlns:m=\"http://purl.oclc.org/ooxml/officeDocument/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://purl.oclc.org/ooxml/wordprocessingml/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" mc:Ignorable=\"w14 w15 w16se wne wp14\"><w:p w:rsidR=\"003217D6\" w:rsidRPr=\"003217D6\" w:rsidRDefault=\"003217D6\" w:rsidP=\"003217D6\"><w:pPr><w:tabs><w:tab w:val=\"end\" w:pos=\"468pt\"/></w:tabs><w:spacing w:line=\"12pt\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:eastAsia=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/><w:b/><w:sz w:val=\"24\"/><w:szCs w:val=\"20\"/></w:rPr></w:pPr><w:r w:rsidRPr=\"003217D6\"><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:eastAsia=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/><w:b/><w:sz w:val=\"24\"/><w:szCs w:val=\"20\"/></w:rPr><w:t>School of Criminal Justice</w:t></w:r><w:r w:rsidRPr=\"003217D6\"><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:eastAsia=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/><w:b/><w:sz w:val=\"24\"/><w:szCs w:val=\"20\"/></w:rPr><w:tab/><w:t xml:space=\"preserve\">   Fall 2016</w:t></w:r></w:p><w:p w:rsidR=\"003217D6\" w:rsidRPr=\"003217D6\" w:rsidRDefault=\"003217D6\" w:rsidP=\"003217D6\"><w:pPr><w:tabs><w:tab w:val=\"end\" w:pos=\"468pt\"/></w:tabs><w:spacing w:line=\"12pt\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:eastAsia=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/><w:b/><w:sz w:val=\"24\"/><w:szCs w:val=\"20\"/></w:rPr></w:pPr><w:r w:rsidRPr=\"003217D6\"><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:eastAsia=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/><w:b/><w:sz w:val=\"24\"/><w:szCs w:val=\"20\"/></w:rPr><w:t xml:space=\"preserve\">CJUS 450</w:t></w:r></w:p><w:p w:rsidR=\"00EA7867\" w:rsidRPr=\"00EA7867\" w:rsidRDefault=\"00835BFA\" w:rsidP=\"00835BFA\"><w:pPr><w:pStyle w:val=\"Header\"/><w:tabs><w:tab w:val=\"clear\" w:pos=\"234pt\"/><w:tab w:val=\"clear\" w:pos=\"468pt\"/><w:tab w:val=\"start\" w:pos=\"346.25pt\"/><w:tab w:val=\"start\" w:pos=\"433.35pt\"/></w:tabs><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/><w:b/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/><w:b/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:tab/></w:r></w:p></w:hdr>");

        file.close();
    }
    public static void writeHeader2File() throws FileNotFoundException
    {
        PrintStream file = new PrintStream(new FileOutputStream("report/word/header2.xml"), false);

        file.printf("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        file.printf("<w:hdr xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://purl.oclc.org/ooxml/officeDocument/relationships\" xmlns:m=\"http://purl.oclc.org/ooxml/officeDocument/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://purl.oclc.org/ooxml/wordprocessingml/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" mc:Ignorable=\"w14 w15 w16se wne wp14\">\n");
        file.printf("    <w:p w:rsidR=\"00DA3F6A\" w:rsidRDefault=\"00DA3F6A\" w:rsidP=\"00DA3F6A\">\n");
        file.printf("        <w:pPr>\n");
        file.printf("            <w:pStyle w:val=\"Header\"/>\n");
        file.printf("            <w:tabs>\n");
        file.printf("                <w:tab w:val=\"clear\" w:pos=\"234pt\"/>\n");
        file.printf("                <w:tab w:val=\"clear\" w:pos=\"468pt\"/>\n");
        file.printf("                <w:tab w:val=\"start\" w:pos=\"433.35pt\"/>\n");
        file.printf("            </w:tabs>\n");
        file.printf("        </w:pPr>\n");
        file.printf("        <w:r>\n");
        file.printf("            <w:rPr>\n");
        file.printf("                <w:noProof/>\n");
        file.printf("            </w:rPr>\n");
        file.printf("            <w:drawing>\n");
        file.printf("                <wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"251659264\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"0\" wp14:anchorId=\"69A817B7\" wp14:editId=\"70A2BD91\">\n");
        file.printf("                    <wp:simplePos x=\"0\" y=\"0\"/>\n");
        file.printf("                    <wp:positionH relativeFrom=\"margin\">\n");
        file.printf("                        <wp:posOffset>1344367</wp:posOffset>\n");
        file.printf("                    </wp:positionH>\n");
        file.printf("                    <wp:positionV relativeFrom=\"paragraph\">\n");
        file.printf("                        <wp:posOffset>0</wp:posOffset>\n");
        file.printf("                    </wp:positionV>\n");
        file.printf("                    <wp:extent cx=\"3582670\" cy=\"548640\"/>\n");
        file.printf("                    <wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"3810\"/>\n");
        file.printf("                    <wp:wrapTopAndBottom/>\n");
        file.printf("                    <wp:docPr id=\"19\" name=\"Picture 19\" descr=\"President LH info 08\"/>\n");
        file.printf("                    <wp:cNvGraphicFramePr>\n");
        file.printf("                        <a:graphicFrameLocks xmlns:a=\"http://purl.oclc.org/ooxml/drawingml/main\" noChangeAspect=\"1\"/>\n");
        file.printf("                    </wp:cNvGraphicFramePr>\n");
        file.printf("                    <a:graphic xmlns:a=\"http://purl.oclc.org/ooxml/drawingml/main\">\n");
        file.printf("                        <a:graphicData uri=\"http://purl.oclc.org/ooxml/drawingml/picture\">\n");
        file.printf("                            <pic:pic xmlns:pic=\"http://purl.oclc.org/ooxml/drawingml/picture\">\n");
        file.printf("                                <pic:nvPicPr>\n");
        file.printf("                                    <pic:cNvPr id=\"0\" name=\"Picture 1\" descr=\"President LH info 08\"/>\n");
        file.printf("                                    <pic:cNvPicPr>\n");
        file.printf("                                        <a:picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\"/>\n");
        file.printf("                                    </pic:cNvPicPr>\n");
        file.printf("                                </pic:nvPicPr>\n");
        file.printf("                                <pic:blipFill>\n");
        file.printf("                                    <a:blip r:embed=\"rId1\">\n");
        file.printf("                                        <a:extLst>\n");
        file.printf("                                            <a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\">\n");
        file.printf("                                                <a14:useLocalDpi xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/>\n");
        file.printf("                                            </a:ext>\n");
        file.printf("                                        </a:extLst>\n");
        file.printf("                                    </a:blip>\n");
        file.printf("                                    <a:srcRect b=\"94.402%%\"/>\n");
        file.printf("                                    <a:stretch>\n");
        file.printf("                                        <a:fillRect/>\n");
        file.printf("                                    </a:stretch>\n");
        file.printf("                                </pic:blipFill>\n");
        file.printf("                                <pic:spPr bwMode=\"auto\">\n");
        file.printf("                                    <a:xfrm>\n");
        file.printf("                                        <a:off x=\"0\" y=\"0\"/>\n");
        file.printf("                                        <a:ext cx=\"3582670\" cy=\"548640\"/>\n");
        file.printf("                                    </a:xfrm>\n");
        file.printf("                                    <a:prstGeom prst=\"rect\">\n");
        file.printf("                                        <a:avLst/>\n");
        file.printf("                                    </a:prstGeom>\n");
        file.printf("                                    <a:noFill/>\n");
        file.printf("                                    <a:ln>\n");
        file.printf("                                        <a:noFill/>\n");
        file.printf("                                    </a:ln>\n");
        file.printf("                                </pic:spPr>\n");
        file.printf("                            </pic:pic>\n");
        file.printf("                        </a:graphicData>\n");
        file.printf("                    </a:graphic>\n");
        file.printf("                    <wp14:sizeRelH relativeFrom=\"page\">\n");
        file.printf("                        <wp14:pctWidth>0%%</wp14:pctWidth>\n");
        file.printf("                    </wp14:sizeRelH>\n");
        file.printf("                    <wp14:sizeRelV relativeFrom=\"page\">\n");
        file.printf("                        <wp14:pctHeight>0%%</wp14:pctHeight>\n");
        file.printf("                    </wp14:sizeRelV>\n");
        file.printf("                </wp:anchor>\n");
        file.printf("            </w:drawing>\n");
        file.printf("        </w:r>\n");
        file.printf("        <w:r>\n");
        file.printf("            <w:tab/>\n");
        file.printf("        </w:r>\n");
        file.printf("    </w:p>\n");
        file.printf("    <w:p w:rsidR=\"00DA3F6A\" w:rsidRPr=\"005E2C4C\" w:rsidRDefault=\"00DA3F6A\" w:rsidP=\"00DA3F6A\">\n");
        file.printf("        <w:pPr>\n");
        file.printf("            <w:tabs>\n");
        file.printf("                <w:tab w:val=\"end\" w:pos=\"468pt\"/>\n");
        file.printf("            </w:tabs>\n");
        file.printf("            <w:spacing w:line=\"12pt\" w:lineRule=\"auto\"/>\n");
        file.printf("            <w:ind w:start=\"36pt\"/>\n");
        file.printf("            <w:jc w:val=\"center\"/>\n");
        file.printf("            <w:rPr>\n");
        file.printf("                <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n");
        file.printf("                <w:b/>\n");
        file.printf("                <w:sz w:val=\"24\"/>\n");
        file.printf("            </w:rPr>\n");
        file.printf("        </w:pPr>\n");
        file.printf("        <w:r>\n");
        file.printf("            <w:rPr>\n");
        file.printf("                <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n");
        file.printf("                <w:b/>\n");
        file.printf("                <w:sz w:val=\"24\"/>\n");
        file.printf("            </w:rPr>\n");
        file.printf("            <w:t>MCOLES Academy Complaint Report System</w:t>\n");
        file.printf("        </w:r>\n");
        file.printf("    </w:p>\n");
        file.printf("    <w:p w:rsidR=\"00DA3F6A\" w:rsidRDefault=\"00DA3F6A\" w:rsidP=\"00DA3F6A\">\n");
        file.printf("        <w:pPr>\n");
        file.printf("            <w:pStyle w:val=\"Header\"/>\n");
        file.printf("            <w:jc w:val=\"center\"/>\n");
        file.printf("            <w:rPr>\n");
        file.printf("                <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n");
        file.printf("                <w:b/>\n");
        file.printf("                <w:sz w:val=\"24\"/>\n");
        file.printf("            </w:rPr>\n");
        file.printf("        </w:pPr>\n");
        file.printf("        <w:r>\n");
        file.printf("            <w:rPr>\n");
        file.printf("                <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n");
        file.printf("                <w:b/>\n");
        file.printf("                <w:sz w:val=\"24\"/>\n");
        file.printf("            </w:rPr>\n");
        file.printf("            <w:t>CJUS 450 Academy Skills Completion</w:t>\n");
        file.printf("        </w:r>\n");
        file.printf("    </w:p>\n");
        file.printf("    <w:p w:rsidR=\"00DA3F6A\" w:rsidRDefault=\"00DA3F6A\" w:rsidP=\"00DA3F6A\">\n");
        file.printf("        <w:pPr>\n");
        file.printf("            <w:pStyle w:val=\"Header\"/>\n");
        file.printf("            <w:jc w:val=\"center\"/>\n");
        file.printf("        </w:pPr>\n");
        file.printf("    </w:p>\n");
        file.printf("</w:hdr>");

        file.close();
    }
 
    public static void writeContentFileStart() throws FileNotFoundException
    {
        input = new PrintStream(new FileOutputStream("report/word/document.xml"), false);
        input.printf("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                     "<w:document xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://purl.oclc.org/ooxml/officeDocument/relationships\" xmlns:m=\"http://purl.oclc.org/ooxml/officeDocument/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://purl.oclc.org/ooxml/wordprocessingml/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" mc:Ignorable=\"w14 w15 w16se wne wp14\" w:conformance=\"strict\">\n" +
                     "    <w:body>\n");
    }

    
    public static void writeBasicInfo(String DOO,
                                      String comp, 
                                      String code,                                      
                                      String OfficeInCharge,
                                      String SecondOfficer,
                                      String Supervisor)
    {
       input.printf("        <w:p w:rsidR=\"00A665B3\" w:rsidRDefault=\"006D52A4\" w:rsidP=\"00E0355A\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:spacing w:line=\"12pt\" w:lineRule=\"auto\"/>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "            <w:bookmarkStart w:id=\"0\" w:name=\"_Hlk479779473\"/>\n" +
                    "            <w:bookmarkEnd w:id=\"0\"/>\n" +
                    "            <w:r>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:noProof/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:drawing>\n" +
                    "                    <wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"251656192\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\" wp14:anchorId=\"3642ACBE\" wp14:editId=\"36E9C8F2\">\n" +
                    "                        <wp:simplePos x=\"0\" y=\"0\"/>\n" +
                    "                        <wp:positionH relativeFrom=\"margin\">\n" +
                    "                            <wp:align>right</wp:align>\n" +
                    "                        </wp:positionH>\n" +
                    "                        <wp:positionV relativeFrom=\"paragraph\">\n" +
                    "                            <wp:posOffset>1905</wp:posOffset>\n" +
                    "                        </wp:positionV>\n" +
                    "                        <wp:extent cx=\"5925820\" cy=\"1129665\"/>\n" +
                    "                        <wp:effectExtent l=\"0\" t=\"0\" r=\"17780\" b=\"13335\"/>\n" +
                    "                        <wp:wrapSquare wrapText=\"bothSides\"/>\n" +
                    "                        <wp:docPr id=\"1\" name=\"Text Box 1\"/>\n" +
                    "                        <wp:cNvGraphicFramePr/>\n" +
                    "                        <a:graphic xmlns:a=\"http://purl.oclc.org/ooxml/drawingml/main\">\n" +
                    "                            <a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">\n" +
                    "                                <wp:wsp>\n" +
                    "                                    <wp:cNvSpPr txBox=\"1\"/>\n" +
                    "                                    <wp:spPr>\n" +
                    "                                        <a:xfrm>\n" +
                    "                                            <a:off x=\"0\" y=\"0\"/>\n" +
                    "                                            <a:ext cx=\"5925820\" cy=\"1129665\"/>\n" +
                    "                                        </a:xfrm>\n" +
                    "                                        <a:prstGeom prst=\"rect\">\n" +
                    "                                            <a:avLst/>\n" +
                    "                                        </a:prstGeom>\n" +
                    "                                        <a:noFill/>\n" +
                    "                                        <a:ln w=\"6350\">\n" +
                    "                                            <a:solidFill>\n" +
                    "                                                <a:prstClr val=\"black\"/>\n" +
                    "                                            </a:solidFill>\n" +
                    "                                        </a:ln>\n" +
                    "                                    </wp:spPr>\n" +
                    "                                    <wp:txbx>\n" +
                    "                                        <wne:txbxContent>\n" +
                    "                                            <w:p w:rsidR=\"006D52A4\" w:rsidRDefault=\"006D52A4\" w:rsidP=\"00E0355A\">\n" +
                    "                                                <w:pPr>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                </w:pPr>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>Date of Offense:</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00E0355A\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">        </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>%10s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                            </w:p>\n" +
                    "                                            <w:p w:rsidR=\"006D52A4\" w:rsidRDefault=\"006D52A4\" w:rsidP=\"00E0355A\">\n" +
                    "                                                <w:pPr>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                </w:pPr>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">Complaint #: </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">            %-10s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                    <w:t xml:space=\"preserve\">    </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                            </w:p>\n" +
                    "                                            <w:p w:rsidR=\"006D52A4\" w:rsidRDefault=\"006D52A4\" w:rsidP=\"00E0355A\">\n" +
                    "                                                <w:pPr>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                </w:pPr>\n" +
                    "                                                <w:r w:rsidRPr=\"00B6309A\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>UCR</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">: </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\"> </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                    <w:t xml:space=\"preserve\">                      %-50s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                            </w:p>\n" +
                    "                                            <w:p w:rsidR=\"006D52A4\" w:rsidRDefault=\"006D52A4\" w:rsidP=\"00E0355A\">\n" +
                    "                                                <w:pPr>\n" +
                    "                                                    <w:spacing w:line=\"12pt\" w:lineRule=\"auto\"/>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                </w:pPr>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">Officer in Charge:    </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">%-50s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                            </w:p>\n" +
                    "                                            <w:p w:rsidR=\"006D52A4\" w:rsidRDefault=\"006D52A4\" w:rsidP=\"00E0355A\">\n" +
                    "                                                <w:pPr>\n" +
                    "                                                    <w:spacing w:line=\"12pt\" w:lineRule=\"auto\"/>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                </w:pPr>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>Secondary Officer:</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">   %-50s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                            </w:p>\n" +
                    "                                            <w:p w:rsidR=\"006D52A4\" w:rsidRDefault=\"006D52A4\" w:rsidP=\"00E0355A\">\n" +
                    "                                                <w:pPr>\n" +
                    "                                                    <w:spacing w:line=\"12pt\" w:lineRule=\"auto\"/>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                </w:pPr>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>Supervising Officer:</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\"> %-50s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                            </w:p>\n" +
                    "                                            <w:p w:rsidR=\"006D52A4\" w:rsidRPr=\"00522587\" w:rsidRDefault=\"006D52A4\" w:rsidP=\"00522587\">\n" +
                    "                                                <w:pPr>\n" +
                    "                                                    <w:spacing w:line=\"12pt\" w:lineRule=\"auto\"/>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                </w:pPr>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">                                                                           </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                            </w:p>\n" +
                    "                                        </wne:txbxContent>\n" +
                    "                                    </wp:txbx>\n" +
                    "                                    <wp:bodyPr rot=\"0\" spcFirstLastPara=\"0\" vertOverflow=\"overflow\" horzOverflow=\"overflow\" vert=\"horz\" wrap=\"square\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" numCol=\"1\" spcCol=\"0\" rtlCol=\"0\" fromWordArt=\"0\" anchor=\"t\" anchorCtr=\"0\" forceAA=\"0\" compatLnSpc=\"1\">\n" +
                    "                                        <a:prstTxWarp prst=\"textNoShape\">\n" +
                    "                                            <a:avLst/>\n" +
                    "                                        </a:prstTxWarp>\n" +
                    "                                        <a:noAutofit/>\n" +
                    "                                    </wp:bodyPr>\n" +
                    "                                </wp:wsp>\n" +
                    "                            </a:graphicData>\n" +
                    "                        </a:graphic>\n" +
                    "                        <wp14:sizeRelH relativeFrom=\"margin\">\n" +
                    "                            <wp14:pctWidth>0%%</wp14:pctWidth>\n" +
                    "                        </wp14:sizeRelH>\n" +
                    "                        <wp14:sizeRelV relativeFrom=\"margin\">\n" +
                    "                            <wp14:pctHeight>0%%</wp14:pctHeight>\n" +
                    "                        </wp14:sizeRelV>\n" +
                    "                    </wp:anchor>\n" +
                    "                </w:drawing>\n" +
                    "            </w:r>\n" +
                    "        </w:p>\n", DOO, comp, code, OfficeInCharge, SecondOfficer, Supervisor);
    }
    
    public static void writeBody(String heading, long postion, int boxNum, int offSet,
                                 String name, String dob, int age, String add1, 
                                 String phone, String add2, String race, String gender, String email)
    {
        System.out.println("In Body");
       input.printf("        <w:p w:rsidR=\"00406F2F\" w:rsidRDefault=\"008B6D3A\" w:rsidP=\"00406F2F\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "            <w:r>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:noProof/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:drawing>\n" +
                    "                    <wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"251657216\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\" wp14:anchorId=\"5E9309CB\" wp14:editId=\"49EB9453\">\n" +
                    "                        <wp:simplePos x=\"0\" y=\"0\"/>\n" +
                    "                        <wp:positionH relativeFrom=\"margin\">\n" +
                    "                            <wp:align>right</wp:align>\n" +
                    "                        </wp:positionH>\n" +
                    "                        <wp:positionV relativeFrom=\"paragraph\">\n" +
                    "                            <wp:posOffset>%d</wp:posOffset>\n" +
                    "                        </wp:positionV>\n" +
                    "                        <wp:extent cx=\"5925820\" cy=\"%d\"/>\n" +
                    "                        <wp:effectExtent l=\"0\" t=\"0\" r=\"17780\" b=\"17145\"/>\n" +
                    "                        <wp:wrapSquare wrapText=\"bothSides\"/>\n" +
                    "                        <wp:docPr id=\"%d\" name=\"Text Box %d\"/>\n" +
                    "                        <wp:cNvGraphicFramePr/>\n" +
                    "                        <a:graphic xmlns:a=\"http://purl.oclc.org/ooxml/drawingml/main\">\n" +
                    "                            <a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">\n" +
                    "                                <wp:wsp>\n" +
                    "                                    <wp:cNvSpPr txBox=\"1\"/>\n" +
                    "                                    <wp:spPr>\n" +
                    "                                        <a:xfrm>\n" +
                    "                                            <a:off x=\"0\" y=\"0\"/>\n" +
                    "                                            <a:ext cx=\"5925820\" cy=\"%d\"/>\n" +
                    "                                        </a:xfrm>\n" +
                    "                                        <a:prstGeom prst=\"rect\">\n" +
                    "                                            <a:avLst/>\n" +
                    "                                        </a:prstGeom>\n" +
                    "                                        <a:noFill/>\n" +
                    "                                        <a:ln w=\"6350\">\n" +
                    "                                            <a:solidFill>\n" +
                    "                                                <a:prstClr val=\"black\"/>\n" +
                    "                                            </a:solidFill>\n" +
                    "                                        </a:ln>\n" +
                    "                                    </wp:spPr>\n" +
                    "                                    <wp:txbx>\n" +
                    "                                        <wne:txbxContent>\n" +
                    "                                            <w:p w:rsidR=\"00406F2F\" w:rsidRPr=\"00B64FA0\" w:rsidRDefault=\"008B6D3A\" w:rsidP=\"00406F2F\">\n" +
                    "                                                <w:pPr>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                </w:pPr>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">Name: </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>%50s      </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                    <w:t xml:space=\"preserve\">DOB: </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>%10s      </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                    <w:t xml:space=\"preserve\">Age: </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>%3d</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                            </w:p>\n" +
                    "                                            <w:p w:rsidR=\"00406F2F\" w:rsidRPr=\"00B64FA0\" w:rsidRDefault=\"008B6D3A\" w:rsidP=\"00406F2F\">\n" +
                    "                                                <w:pPr>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                </w:pPr>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">%-30s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">                         </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                    <w:t>Phone:</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\"> </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>%-14s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                            </w:p>\n" +
                    "                                            <w:p w:rsidR=\"00406F2F\" w:rsidRPr=\"00B64FA0\" w:rsidRDefault=\"008B6D3A\" w:rsidP=\"00406F2F\">\n" +
                    "                                                <w:pPr>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                </w:pPr>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>%50s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                    <w:t xml:space=\"preserve\">    </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                            </w:p>\n" +
                    "                                            <w:p w:rsidR=\"00406F2F\" w:rsidRPr=\"00B64FA0\" w:rsidRDefault=\"008B6D3A\" w:rsidP=\"00406F2F\">\n" +
                    "                                                <w:pPr>\n" +
                    "                                                    <w:spacing w:line=\"12pt\" w:lineRule=\"auto\"/>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                </w:pPr>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">Race: </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>%s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                    <w:t>Sex:</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\"> </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>%6s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                    <w:t xml:space=\"preserve\">Email: %30s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                            </w:p>\n" +
                    "                                        </wne:txbxContent>\n" +
                    "                                    </wp:txbx>\n" +
                    "                                    <wp:bodyPr rot=\"0\" spcFirstLastPara=\"0\" vertOverflow=\"overflow\" horzOverflow=\"overflow\" vert=\"horz\" wrap=\"square\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" numCol=\"1\" spcCol=\"0\" rtlCol=\"0\" fromWordArt=\"0\" anchor=\"t\" anchorCtr=\"0\" forceAA=\"0\" compatLnSpc=\"1\">\n" +
                    "                                        <a:prstTxWarp prst=\"textNoShape\">\n" +
                    "                                            <a:avLst/>\n" +
                    "                                        </a:prstTxWarp>\n" +
                    "                                        <a:noAutofit/>\n" +
                    "                                    </wp:bodyPr>\n" +
                    "                                </wp:wsp>\n" +
                    "                            </a:graphicData>\n" +
                    "                        </a:graphic>\n" +
                    "                        <wp14:sizeRelH relativeFrom=\"margin\">\n" +
                    "                            <wp14:pctWidth>0%%</wp14:pctWidth>\n" +
                    "                        </wp14:sizeRelH>\n" +
                    "                        <wp14:sizeRelV relativeFrom=\"margin\">\n" +
                    "                            <wp14:pctHeight>0%%</wp14:pctHeight>\n" +
                    "                        </wp14:sizeRelV>\n" +
                    "                    </wp:anchor>\n" +
                    "                </w:drawing>\n" +
                    "            </w:r>\n" +
                    "            <w:r>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t>%30s</w:t>\n" +
                    "            </w:r>\n" +
                    "        </w:p>\n" +
                    "        <w:p w:rsidR=\"00F560FD\" w:rsidRPr=\"00A665B3\" w:rsidRDefault=\"00F560FD\" w:rsidP=\"00A665B3\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "        </w:p>\n", offSet,postion, boxNum, boxNum, postion, name, dob, age, add1, phone, add2, race, gender, email, heading);
    }
    
    public static void writeEvidence(String heading, long postion, int boxNum, int offSet,
                                     String name, String dob, int age, String add1, 
                                     String phone, String add2, String race, String gender, String email)
    {
        System.out.println("In Body");
       input.printf("        <w:p w:rsidR=\"00406F2F\" w:rsidRDefault=\"008B6D3A\" w:rsidP=\"00406F2F\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "            <w:r>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:noProof/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:drawing>\n" +
                    "                    <wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"251657216\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\" wp14:anchorId=\"5E9309CB\" wp14:editId=\"49EB9453\">\n" +
                    "                        <wp:simplePos x=\"0\" y=\"0\"/>\n" +
                    "                        <wp:positionH relativeFrom=\"margin\">\n" +
                    "                            <wp:align>right</wp:align>\n" +
                    "                        </wp:positionH>\n" +
                    "                        <wp:positionV relativeFrom=\"paragraph\">\n" +
                    "                            <wp:posOffset>%d</wp:posOffset>\n" +
                    "                        </wp:positionV>\n" +
                    "                        <wp:extent cx=\"5925820\" cy=\"%d\"/>\n" +
                    "                        <wp:effectExtent l=\"0\" t=\"0\" r=\"17780\" b=\"17145\"/>\n" +
                    "                        <wp:wrapSquare wrapText=\"bothSides\"/>\n" +
                    "                        <wp:docPr id=\"%d\" name=\"Text Box %d\"/>\n" +
                    "                        <wp:cNvGraphicFramePr/>\n" +
                    "                        <a:graphic xmlns:a=\"http://purl.oclc.org/ooxml/drawingml/main\">\n" +
                    "                            <a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">\n" +
                    "                                <wp:wsp>\n" +
                    "                                    <wp:cNvSpPr txBox=\"1\"/>\n" +
                    "                                    <wp:spPr>\n" +
                    "                                        <a:xfrm>\n" +
                    "                                            <a:off x=\"0\" y=\"0\"/>\n" +
                    "                                            <a:ext cx=\"5925820\" cy=\"%d\"/>\n" +
                    "                                        </a:xfrm>\n" +
                    "                                        <a:prstGeom prst=\"rect\">\n" +
                    "                                            <a:avLst/>\n" +
                    "                                        </a:prstGeom>\n" +
                    "                                        <a:noFill/>\n" +
                    "                                        <a:ln w=\"6350\">\n" +
                    "                                            <a:solidFill>\n" +
                    "                                                <a:prstClr val=\"black\"/>\n" +
                    "                                            </a:solidFill>\n" +
                    "                                        </a:ln>\n" +
                    "                                    </wp:spPr>\n" +
                    "                                    <wp:txbx>\n" +
                    "                                        <wne:txbxContent>\n" +
                    "                                            <w:p w:rsidR=\"00406F2F\" w:rsidRPr=\"00B64FA0\" w:rsidRDefault=\"008B6D3A\" w:rsidP=\"00406F2F\">\n" +
                    "                                                <w:pPr>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                </w:pPr>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">Name: </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>%50s      </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                    <w:t xml:space=\"preserve\">DOB: </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>%10s      </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                    <w:t xml:space=\"preserve\">Age: </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>%3d</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                            </w:p>\n" +
                    "                                            <w:p w:rsidR=\"00406F2F\" w:rsidRPr=\"00B64FA0\" w:rsidRDefault=\"008B6D3A\" w:rsidP=\"00406F2F\">\n" +
                    "                                                <w:pPr>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                </w:pPr>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">%-30s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">                         </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                    <w:t>Phone:</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\"> </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>%-14s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                            </w:p>\n" +
                    "                                            <w:p w:rsidR=\"00406F2F\" w:rsidRPr=\"00B64FA0\" w:rsidRDefault=\"008B6D3A\" w:rsidP=\"00406F2F\">\n" +
                    "                                                <w:pPr>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                </w:pPr>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>%50s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                    <w:t xml:space=\"preserve\">    </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                            </w:p>\n" +
                    "                                            <w:p w:rsidR=\"00406F2F\" w:rsidRPr=\"00B64FA0\" w:rsidRDefault=\"008B6D3A\" w:rsidP=\"00406F2F\">\n" +
                    "                                                <w:pPr>\n" +
                    "                                                    <w:spacing w:line=\"12pt\" w:lineRule=\"auto\"/>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                </w:pPr>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\">Race: </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>%s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                    <w:t>Sex:</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r>\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t xml:space=\"preserve\"> </w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                        <w:b/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:t>%6s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                                <w:r w:rsidRPr=\"00B64FA0\">\n" +
                    "                                                    <w:rPr>\n" +
                    "                                                        <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                                                    </w:rPr>\n" +
                    "                                                    <w:tab/>\n" +
                    "                                                    <w:t xml:space=\"preserve\">Email: %30s</w:t>\n" +
                    "                                                </w:r>\n" +
                    "                                            </w:p>\n" +
                    "                                        </wne:txbxContent>\n" +
                    "                                    </wp:txbx>\n" +
                    "                                    <wp:bodyPr rot=\"0\" spcFirstLastPara=\"0\" vertOverflow=\"overflow\" horzOverflow=\"overflow\" vert=\"horz\" wrap=\"square\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" numCol=\"1\" spcCol=\"0\" rtlCol=\"0\" fromWordArt=\"0\" anchor=\"t\" anchorCtr=\"0\" forceAA=\"0\" compatLnSpc=\"1\">\n" +
                    "                                        <a:prstTxWarp prst=\"textNoShape\">\n" +
                    "                                            <a:avLst/>\n" +
                    "                                        </a:prstTxWarp>\n" +
                    "                                        <a:noAutofit/>\n" +
                    "                                    </wp:bodyPr>\n" +
                    "                                </wp:wsp>\n" +
                    "                            </a:graphicData>\n" +
                    "                        </a:graphic>\n" +
                    "                        <wp14:sizeRelH relativeFrom=\"margin\">\n" +
                    "                            <wp14:pctWidth>0%%</wp14:pctWidth>\n" +
                    "                        </wp14:sizeRelH>\n" +
                    "                        <wp14:sizeRelV relativeFrom=\"margin\">\n" +
                    "                            <wp14:pctHeight>0%%</wp14:pctHeight>\n" +
                    "                        </wp14:sizeRelV>\n" +
                    "                    </wp:anchor>\n" +
                    "                </w:drawing>\n" +
                    "            </w:r>\n" +
                    "            <w:r>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t>%30s</w:t>\n" +
                    "            </w:r>\n" +
                    "        </w:p>\n" +
                    "        <w:p w:rsidR=\"00F560FD\" w:rsidRPr=\"00A665B3\" w:rsidRDefault=\"00F560FD\" w:rsidP=\"00A665B3\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "        </w:p>\n", offSet,postion, boxNum, boxNum, postion, name, dob, age, add1, phone, add2, race, gender, email, heading);
    }
    
    public static void sectionDivider(String ReportingOfficer, String dateOfReport)
    {
       input.printf("        <w:p w:rsidR=\"00F560FD\" w:rsidRDefault=\"0030196B\" w:rsidP=\"00A665B3\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "            <w:r>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t>_____________________________________________________________________________________</w:t>\n" +
                    "            </w:r>\n" +
                    "        </w:p>\n" +
                    "        <w:p w:rsidR=\"0030196B\" w:rsidRDefault=\"009F1C92\" w:rsidP=\"00A665B3\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:pBdr>\n" +
                    "                    <w:bottom w:val=\"single\" w:sz=\"6\" w:space=\"1\" w:color=\"auto\"/>\n" +
                    "                </w:pBdr>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:sz w:val=\"20\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "            <w:r w:rsidRPr=\"009F1C92\">\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:sz w:val=\"20\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t>Reporting Officer:</w:t>\n" +
                    "            </w:r>\n" +
                    "            <w:r>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:sz w:val=\"20\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t xml:space=\"preserve\"> %s </w:t>\n" +
                    "            </w:r>\n" +
                    "            <w:r>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:sz w:val=\"20\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:tab/>\n" +
                    "            </w:r>\n" +
                    "            <w:r>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:sz w:val=\"20\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:tab/>\n" +
                    "                <w:t xml:space=\"preserve\">          </w:t>\n" +
                    "            </w:r>\n" +
                    "            <w:r w:rsidR=\"003217D6\">\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:sz w:val=\"20\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:tab/>\n" +
                    "            </w:r>\n" +
                    "            <w:r w:rsidR=\"003217D6\">\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:sz w:val=\"20\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:tab/>\n" +
                    "            </w:r>\n" +
                    "            <w:r w:rsidR=\"003217D6\">\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:sz w:val=\"20\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:tab/>\n" +
                    "            </w:r>\n" +
                    "            <w:r w:rsidR=\"003217D6\">\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:sz w:val=\"20\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:tab/>\n" +
                    "            </w:r>\n" +
                    "            <w:r w:rsidR=\"003217D6\">\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:sz w:val=\"20\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:tab/>\n" +
                    "            </w:r>\n" +
                    "            <w:r w:rsidR=\"003217D6\">\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:sz w:val=\"20\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:tab/>\n" +
                    "                <w:t xml:space=\"preserve\">          </w:t>\n" +
                    "            </w:r>\n" +
                    "            <w:r>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:sz w:val=\"20\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t>%10s</w:t>\n" +
                    "            </w:r>\n" +
                    "        </w:p>\n", ReportingOfficer, dateOfReport);
    }
    
    public static void writeInformation(String heading)
    {
        System.out.println("In Body");
       input.printf("        <w:p w:rsidR=\"00D9192E\" w:rsidRPr=\"00D9192E\" w:rsidRDefault=\"00D9192E\" w:rsidP=\"00D9192E\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:tabs>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"-36pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"0pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"15.05pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"36pt\"/>\n" +
                    "                </w:tabs>\n" +
                    "                <w:suppressAutoHyphens/>\n" +
                    "                <w:ind w:start=\"15.05pt\" w:end=\"15.05pt\" w:hanging=\"15.05pt\"/>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "            <w:r w:rsidRPr=\"00D9192E\">\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t>%s</w:t>\n" +
                    "            </w:r>\n" +
                    "        </w:p>\n" +
                    "        <w:p w:rsidR=\"00461B02\" w:rsidRPr=\"00D9192E\" w:rsidRDefault=\"00D9192E\" w:rsidP=\"00BE7629\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:tabs>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"-36pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"0pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"15.05pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"36pt\"/>\n" +
                    "                </w:tabs>\n" +
                    "                <w:suppressAutoHyphens/>\n" +
                    "                <w:ind w:start=\"15.05pt\" w:end=\"15.05pt\" w:hanging=\"15.05pt\"/>\n"   +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "        </w:p>\n" +
                    "        <w:p w:rsidR=\"00461B02\" w:rsidRDefault=\"00461B02\" w:rsidP=\"00461B02\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:tabs>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"-36pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"0pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"15.05pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"36pt\"/>\n" +
                    "                </w:tabs>\n" +
                    "                <w:suppressAutoHyphens/>\n" +
                    "                <w:ind w:start=\"15.05pt\" w:end=\"15.05pt\" w:hanging=\"15.05pt\"/>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "            <w:r>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t>Type here</w:t>\n" +
                    "            </w:r>\n" +
                    "        </w:p>\n" +
                    "        <w:p w:rsidR=\"00D26F80\" w:rsidRDefault=\"00D26F80\" w:rsidP=\"00461B02\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:tabs>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"-36pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"0pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"15.05pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"36pt\"/>\n" +
                    "                </w:tabs>\n" +
                    "                <w:suppressAutoHyphens/>\n" +
                    "                <w:ind w:start=\"15.05pt\" w:end=\"15.05pt\" w:hanging=\"15.05pt\"/>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "        </w:p>\n", heading);
    }
    
    public static void writeHeadings(String heading, String name)
    {
       input.printf("        <w:p w:rsidR=\"00D9192E\" w:rsidRPr=\"00D9192E\" w:rsidRDefault=\"00D9192E\" w:rsidP=\"00461B02\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:tabs>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"-36pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"0pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"15.05pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"36pt\"/>\n" +
                    "                </w:tabs>\n" +
                    "                <w:suppressAutoHyphens/>\n" +
                    "                <w:ind w:start=\"15.05pt\" w:end=\"15.05pt\" w:hanging=\"15.05pt\"/>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "            <w:r w:rsidRPr=\"00D9192E\">\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t>INTERVIEW OF %s (</w:t>\n" +
                    "            </w:r>\n" +
                    "            <w:r w:rsidR=\"00461B02\">\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t>%s</w:t>\n" +
                    "            </w:r>\n" +
                    "            <w:r w:rsidRPr=\"00D9192E\">\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t>)</w:t>\n" +
                    "            </w:r>\n" +
                    "        </w:p>\n" +
                    "        <w:p w:rsidR=\"00D9192E\" w:rsidRDefault=\"00D9192E\" w:rsidP=\"00461B02\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:tabs>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"-36pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"0pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"15.05pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"36pt\"/>\n" +
                    "                </w:tabs>\n" +
                    "                <w:suppressAutoHyphens/>\n" +
                    "                <w:ind w:start=\"15.05pt\" w:end=\"15.05pt\" w:hanging=\"15.05pt\"/>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "        </w:p>\n" +
                    "        <w:p w:rsidR=\"00BE7629\" w:rsidRPr=\"00D9192E\" w:rsidRDefault=\"00786868\" w:rsidP=\"00BE7629\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:tabs>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"-36pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"0pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"36pt\"/>\n" +
                    "                </w:tabs>\n" +
                    "                <w:suppressAutoHyphens/>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "            <w:r>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t>Type here</w:t>\n" +
                    "            </w:r>\n" +
                    "        </w:p>\n" +
                    "        <w:p w:rsidR=\"00BE7629\" w:rsidRPr=\"00D9192E\" w:rsidRDefault=\"00BE7629\" w:rsidP=\"00BE7629\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:tabs>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"-36pt\"/>\n" +
                    "                </w:tabs>\n" +
                    "                <w:suppressAutoHyphens/>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "        </w:p>\n", heading, name);
    }
    
    public static void writeContentFileEnd() throws FileNotFoundException
    {
       input.printf("        <w:p w:rsidR=\"00D9192E\" w:rsidRPr=\"00D9192E\" w:rsidRDefault=\"00D9192E\" w:rsidP=\"00DA3660\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:tabs>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"-36pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"0pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"36pt\"/>\n" +
                    "                </w:tabs>\n" +
                    "                <w:suppressAutoHyphens/>\n" +
                    "                <w:ind w:end=\"-71.55pt\"/>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "            <w:r w:rsidRPr=\"00D9192E\">\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t>OFFICER SIGNATURE   _________________</w:t>\n" +
                    "            </w:r>\n" +
                    "            <w:r w:rsidR=\"00DA3660\">\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t xml:space=\"preserve\">__________________________ DATE </w:t>\n" +
                    "            </w:r>\n" +
                    "            <w:r w:rsidRPr=\"00D9192E\">\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t>_____________</w:t>\n" +
                    "            </w:r>\n" +
                    "        </w:p>\n" +
                    "        <w:p w:rsidR=\"00D9192E\" w:rsidRPr=\"00D9192E\" w:rsidRDefault=\"00D9192E\" w:rsidP=\"00461B02\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:tabs>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"-36pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"0pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"15.05pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"36pt\"/>\n" +
                    "                </w:tabs>\n" +
                    "                <w:suppressAutoHyphens/>\n" +
                    "                <w:ind w:start=\"15.05pt\" w:end=\"15.05pt\" w:hanging=\"15.05pt\"/>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "        </w:p>\n" +
                    "        <w:p w:rsidR=\"00D9192E\" w:rsidRPr=\"00D9192E\" w:rsidRDefault=\"00D9192E\" w:rsidP=\"00DA3660\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:tabs>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"-36pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"0pt\"/>\n" +
                    "                    <w:tab w:val=\"start\" w:pos=\"36pt\"/>\n" +
                    "                </w:tabs>\n" +
                    "                <w:suppressAutoHyphens/>\n" +
                    "                <w:ind w:end=\"-72pt\"/>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "            <w:r w:rsidRPr=\"00D9192E\">\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>\n" +
                    "                    <w:b/>\n" +
                    "                    <w:szCs w:val=\"24\"/>\n" +
                    "                    <w:u w:val=\"single\"/>\n" +
                    "                </w:rPr>\n" +
                    "                <w:t>SUPERVISOR SIGNATURE _________________________________________DATE _____________</w:t>\n" +
                    "            </w:r>\n" +
                    "        </w:p>\n" +
                    "        <w:p w:rsidR=\"00835BFA\" w:rsidRPr=\"00D9192E\" w:rsidRDefault=\"00835BFA\" w:rsidP=\"00461B02\">\n" +
                    "            <w:pPr>\n" +
                    "                <w:spacing w:line=\"12pt\" w:lineRule=\"auto\"/>\n" +
                    "                <w:rPr>\n" +
                    "                    <w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>\n" +
                    "                    <w:sz w:val=\"20\"/>\n" +
                    "                </w:rPr>\n" +
                    "            </w:pPr>\n" +
                    "        </w:p>\n" +
                    "        <w:sectPr w:rsidR=\"00835BFA\" w:rsidRPr=\"00D9192E\" w:rsidSect=\"00DA3F6A\">\n" +
                    "            <w:headerReference w:type=\"default\" r:id=\"rId7\"/>\n" +
                    "            <w:footerReference w:type=\"default\" r:id=\"rId8\"/>\n" +
                    "            <w:headerReference w:type=\"first\" r:id=\"rId9\"/>\n" +
                    "            <w:footerReference w:type=\"first\" r:id=\"rId10\"/>\n" +
                    "            <w:pgSz w:w=\"612pt\" w:h=\"792pt\"/>\n" +
                    "            <w:pgMar w:top=\"72pt\" w:right=\"72pt\" w:bottom=\"72pt\" w:left=\"72pt\" w:header=\"36pt\" w:footer=\"36pt\" w:gutter=\"0pt\"/>\n" +
                    "            <w:cols w:space=\"36pt\"/>\n" +
                    "            <w:titlePg/>\n" +
                    "            <w:docGrid w:linePitch=\"360\"/>\n" +
                    "        </w:sectPr>\n" +
                    "    </w:body>\n" +
                    "</w:document>");
        input.close();
    }
	
    public static void createDOCXArchive() throws IOException
    {
        LocalDateTime now = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd-HH-mm-ss");

        String formatDateTime = now.format(formatter);
        
        // now = LocalDateTime.parse(date, formatter);
       
        String OS;
        OS = System.getProperty("os.name");
        if (OS.startsWith("Windows"))
        {
            String oog;
            Process cmdProc = Runtime.getRuntime().exec("7z a -tzip -r ./report-" + formatDateTime + ".docx ./report/*");
            BufferedReader stdoutReader = new BufferedReader(new InputStreamReader(cmdProc.getInputStream()));
            while ((oog = stdoutReader.readLine()) != null) 
            {
                System.out.printf("%s\n", oog);       // process procs standard output here
            }

            BufferedReader stderrReader = new BufferedReader(new InputStreamReader(cmdProc.getErrorStream()));
            while ((stderrReader.readLine()) != null) 
            {
                // process procs standard error here
            }

            cmdProc.exitValue();
        }
        else
        {
            String oog;
            Process cmdProc = Runtime.getRuntime().exec("zip -r report.docx .", null, new File("./report"));
            BufferedReader stdoutReader = new BufferedReader(new InputStreamReader(cmdProc.getInputStream()));
            while ((oog =stdoutReader.readLine()) != null) 
            {
                System.out.printf("%s\n", oog);
			   // process procs standard output here
            }

            BufferedReader stderrReader = new BufferedReader(new InputStreamReader(cmdProc.getErrorStream()));
            while ((stderrReader.readLine()) != null) 
            {
			   // process procs standard error here
            }
            try
            {
                cmdProc.waitFor();
            } catch (Exception ex)
            {}
            System.out.printf("Value is %d\n", cmdProc.exitValue());
        }
    } 
    
    public static void checkFor7z() throws IOException
    {
        String OS;
        OS = System.getProperty("os.name");
        if (OS.startsWith("Windows"))
        {
            String oog;
            Process cmdProc = Runtime.getRuntime().exec("7z");
            BufferedReader stdoutReader = new BufferedReader(new InputStreamReader(cmdProc.getInputStream()));
            while ((oog = stdoutReader.readLine()) != null) 
            {
                System.out.printf("%s\n", oog);       // process procs standard output here
            }

            BufferedReader stderrReader = new BufferedReader(new InputStreamReader(cmdProc.getErrorStream()));
            while ((stderrReader.readLine()) != null) 
            {
                // process procs standard error here
            }

            cmdProc.exitValue();
        }
    }

    public static void writeSettingsFile() throws FileNotFoundException
    {
        PrintStream file = new PrintStream(new FileOutputStream("report/word/settings.xml"), false);

        file.print("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                    "<w:settings xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://purl.oclc.org/ooxml/officeDocument/relationships\" xmlns:m=\"http://purl.oclc.org/ooxml/officeDocument/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://purl.oclc.org/ooxml/wordprocessingml/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:sl=\"http://schemas.openxmlformats.org/schemaLibrary/2006/main\" mc:Ignorable=\"w14 w15 w16se\"><w:zoom w:percent=\"110%\"/><w:proofState w:spelling=\"clean\" w:grammar=\"clean\"/><w:defaultTabStop w:val=\"36pt\"/><w:characterSpacingControl w:val=\"doNotCompress\"/><w:footnotePr><w:footnote w:id=\"-1\"/><w:footnote w:id=\"0\"/></w:footnotePr><w:endnotePr><w:endnote w:id=\"-1\"/><w:endnote w:id=\"0\"/></w:endnotePr><w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"15\"/><w:compatSetting w:name=\"overrideTableStyleFontSizeAndJustification\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/><w:compatSetting w:name=\"enableOpenTypeFeatures\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/><w:compatSetting w:name=\"doNotFlipMirrorIndents\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/><w:compatSetting w:name=\"differentiateMultirowTableHeaders\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/></w:compat><w:rsids><w:rsidRoot w:val=\"004A0575\"/><w:rsid w:val=\"00232037\"/><w:rsid w:val=\"003A11A1\"/><w:rsid w:val=\"004A0575\"/><w:rsid w:val=\"006D52A4\"/><w:rsid w:val=\"00B6309A\"/><w:rsid w:val=\"00C2303C\"/><w:rsid w:val=\"00E0355A\"/><w:rsid w:val=\"00E72372\"/><w:rsid w:val=\"00EA7867\"/><w:rsid w:val=\"00EF7AA9\"/><w:rsid w:val=\"00F21BF0\"/><w:rsid w:val=\"00F560FD\"/></w:rsids><m:mathPr><m:mathFont m:val=\"Cambria Math\"/><m:brkBin m:val=\"before\"/><m:brkBinSub m:val=\"--\"/><m:smallFrac m:val=\"0\"/><m:dispDef/><m:lMargin m:val=\"0\"/><m:rMargin m:val=\"0\"/><m:defJc m:val=\"centerGroup\"/><m:wrapIndent m:val=\"1440\"/><m:intLim m:val=\"subSup\"/><m:naryLim m:val=\"undOvr\"/></m:mathPr><w:themeFontLang w:val=\"en-US\"/><w:clrSchemeMapping w:bg1=\"light1\" w:t1=\"dark1\" w:bg2=\"light2\" w:t2=\"dark2\" w:accent1=\"accent1\" w:accent2=\"accent2\" w:accent3=\"accent3\" w:accent4=\"accent4\" w:accent5=\"accent5\" w:accent6=\"accent6\" w:hyperlink=\"hyperlink\" w:followedHyperlink=\"followedHyperlink\"/><w:decimalSymbol w:val=\".\"/><w:listSeparator w:val=\",\"/><w14:docId w14:val=\"32969A88\"/><w15:chartTrackingRefBased/><w15:docId w15:val=\"{6A362AC2-9494-4C0C-8611-C6E8FC0B894B}\"/></w:settings>");

        file.close();
    }

    public static void writeStylesFile()
    {
//        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }
}

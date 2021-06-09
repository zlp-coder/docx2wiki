package op.tools.docx2wiki;

import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.UUID;


public class DocTransfer {

    public static Logger log = LoggerFactory.getLogger(DocTransfer.class);

    public void Transfer(String fileName, FileInputStream FileStream, WikiOperator mOp) {
        log.info("按 DOCX 文档格式开始进行处理文档...");

        XWPFDocument doc = null;
        try {
            doc = new XWPFDocument(FileStream);
        } catch (Exception ex) {
            log.error("无法在加载文件，Error：" + ex.getMessage());
            return;
        }

        try {
            String strTable = "";
            String sHead = "==标签==\r\n  " + fileName.substring(0, fileName.lastIndexOf(".")) + " ";  //公共关联
            String sTitle = "";
            StringBuilder sOutIndex = new StringBuilder();

            String sTileAll = fileName.substring(0, fileName.lastIndexOf("."));

            StringBuilder sOut = new StringBuilder();
            StringBuilder sOutAll = new StringBuilder();

            sOutIndex.append("==[[" + sTileAll + "]]==\r");

            //图形文件的缓存路径
            String BMPPath = new PathTools().getPath() + File.separator + "temp" + File.separator;
            File fbmp = new File(BMPPath);

            if (!fbmp.exists()) {
                fbmp.mkdir();
            }

            //region>>扫描Word，转换格式
            for (int i = 1; i < doc.getBodyElements().size(); i++) //word逐段扫描
            {
                BodyElementType elementStyle = doc.getBodyElements().get(i).getElementType();

                if (elementStyle == BodyElementType.PARAGRAPH) {
                    XWPFParagraph p = (XWPFParagraph) doc.getBodyElements().get(i);
                    String localStyle = p.getStyle();

                    if (localStyle != null && localStyle.equals("1")) //H1，新开页面写入
                    {
                        if (p.getText().isEmpty() == true) {
                            continue;
                        }

                        sOutAll.append("==" + FormatString(p.getText()) + "==\r");
                        continue;
                    }

                    if (localStyle != null && localStyle.equals("2")) //H2
                    {
                        if (p.getText().isEmpty() == true) {
                            continue;
                        }

                        sOutAll.append("===" + FormatString(p.getText()) + "===\r");
                        continue;
                    }

                    if (localStyle != null && localStyle.equals("3")) //H3
                    {
                        if (p.getText().isEmpty() == true) {
                            continue;
                        }

                        sOutAll.append("====" + FormatString(p.getText()) + "====\r");
                        continue;
                    }

                    if (localStyle != null && localStyle.equals("4")) {
                        if (p.getText().isEmpty() == true) {
                            continue;
                        }

                        sOutAll.append("=====" + FormatString(p.getText()) + "=====\r");
                        continue;
                    }

                    if (p.getText().isEmpty() != true) //其他都作为正文处理
                    {
                        if (!p.getText().isEmpty()) {
                            if (p.getNumLevelText() != null) {
                                if (p.getNumLevelText().indexOf("1") >= 0) {
                                    sOutAll.append("#");
                                }
                                if (p.getNumLevelText().indexOf("2") >= 0) {
                                    sOutAll.append("##");
                                }
                                if (p.getNumLevelText().indexOf("3") >= 0) {
                                    sOutAll.append("###");
                                }
                            }
                            sOutAll.append("" + FormatString(p.getText()) + "<br/>\r");
                        }
                    }

                    if (p.getRuns().size() > 0) {
                        for (XWPFRun r : p.getRuns()) {
                            if (r.getEmbeddedPictures().size() > 0) {
                                for (XWPFPicture xwpfpic : r.getEmbeddedPictures()) {
                                    XWPFPictureData pic = xwpfpic.getPictureData();
                                    log.info("捕获到嵌入图形，开始处理" + pic.getFileName() + "...");

                                    //String fName = sTileAll + "_" + pic.getFileName();

                                    String fName = UUID.randomUUID() + "_" + pic.getFileName();
                                    UploadBmpInfo info = new UploadBmpInfo();
                                    info.set_FileName(fName);
                                    //info.set_FileFullName(new PathTools().getPath() + File.separator + "temp" +
                                    //File.separator + fName);

                                    SimpleDateFormat sdfDate = new SimpleDateFormat("yyyy-MM-dd hh:mm");
                                    info.set_Date(sdfDate.format(new Date()));

                                    info.set_Description("Doc2Wiki程序自动上传的图片");
                                    info.set_Source(fileName);
                                    info.set_BmpData(pic.getData());

                                    //info.set_FileName(info.get_FileName().replace(".emf", ".png"));
//                                    info.set_FileFullName(info.get_FileFullName().replace(".emf", ".png"));


                                    mOp.addBmpInfo(info);
                                    sOutAll.append("[[Image:" + info.get_FileName().replace(".emf", ".png") + "]]\r\r");
                                }
                            }
                        }
                    }

                }

                if (elementStyle == BodyElementType.TABLE) {
                    XWPFTable t = (XWPFTable) doc.getBodyElements().get(i);

                    StringBuilder tmpTable = new StringBuilder();
                    tmpTable.append("<table class='wikitable'>");

                    String blockMark = "";

                    for (int irow = 0; irow < t.getRows().size(); irow++) {

                        XWPFTableRow row = t.getRow(irow);
                        tmpTable.append("<tr style='height=25'>");

                        for (int icell = 0; icell < row.getTableCells().size(); icell++) {
                            XWPFTableCell cell = row.getCell(icell);

                            String tmpNote = cell.getTextRecursively().isEmpty() ? "&nbsp;" : cell.getTextRecursively();
                            tmpNote = tmpNote.replace("\t", "<BR/>");

                            Integer colspan = 0;
                            if (cell.getCTTc().getTcPr().getGridSpan() != null) {
                                long span = cell.getCTTc().getTcPr().getGridSpan().getVal().longValue();
                                colspan = Integer.parseInt(Long.toString(span));
                            }

                            Integer rowspan = 0;
                            boolean parseThiscell = false;

                            if( blockMark.indexOf("{" + irow + "|" +icell + "}") < 0  ) {
                                for (int j = irow +1; j <= t.getRows().size(); j++) {
                                    if (t.getRow(j) != null) {
                                        XWPFTableCell nextCell = t.getRow(j).getCell(icell);

                                        if (nextCell != null && nextCell.getCTTc().getTcPr().getVMerge() != null) {
                                            rowspan = rowspan + 1;

                                            blockMark += "{" + j + "|" +icell + "}";
                                        } else {
                                            break;
                                        }
                                    }
                                }
                            }else {
                                parseThiscell = true;
                            }

                            if(parseThiscell ==false) {
                                tmpTable.append("<td");
                                if (colspan != null && colspan > 0) {
                                    tmpTable.append(" colspan=" + colspan);
                                }
                                if (rowspan > 0) {
                                    tmpTable.append(" rowspan=" + (rowspan +1));
                                }
                                tmpTable.append(">");

                                tmpTable.append(tmpNote);
                                tmpTable.append("</td>");
                            }

                        }
                        tmpTable.append("</tr>");
                    }

                    tmpTable.append("</table>");
                    sOutAll.append(tmpTable.toString() + "\r");
                }
            }

            mOp.set_title(sTileAll);
            mOp.set_text(sOutAll.toString());

            log.info("写入新页面：" + sTileAll);

        } catch (Exception ex) {
            ex.printStackTrace();
            log.error("Error：" + fileName + "处理失败。需要重新处理。" + ex.getMessage());
        }
    }

    protected String FormatString(String sData) {
        String sR = sData.trim();
        sR = sR.replace("\r\n", "<BR/>");
        sR = sR.replace("\r", "<BR/>");
        sR = sR.replace("\t", "&nbsp;");
        return sR;
    }

    protected String FormatTable(String sData) {
        String sR = "";
        sR += "{|class=\"wikitable\" border=\"1\" \r|";
        sR += sData;
        sR = sR.replace("\r\n", "\r|-\r|");
        sR = sR.replace("\t", "\r|");
        sR = sR.substring(0, sR.length() - 2);
        sR += "\r|}";

        return sR;
    }

}

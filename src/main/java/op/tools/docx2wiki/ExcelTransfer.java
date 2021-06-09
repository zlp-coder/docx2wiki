package op.tools.docx2wiki;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;

public class ExcelTransfer {
    public static Logger log = LoggerFactory.getLogger(DocTransfer.class);

    public String wikiUserName = "";

    public void Transfer(String fileName, FileInputStream fis, WikiOperator op) {
        log.info("按 EXCEL 文档格式开始进行处理文档...");

        HSSFWorkbook workbook = null;
        try {
            workbook = new HSSFWorkbook(fis);
        } catch (Exception ex) {
            log.error("无法在 NPIO 中加载文件，Error：" + ex.getMessage());
            return;
        }

        try {
            String strTable = "";
            String sHead = "==标签==\r\n  " + fileName.substring(0, fileName.lastIndexOf(".")) + " ";
            String sTitle = "";
            StringBuilder sOutIndex = new StringBuilder();
            String sTileAll = fileName.substring(0, fileName.lastIndexOf("."));

            StringBuilder sOut = new StringBuilder();
            StringBuilder sOutAll = new StringBuilder();

            sOutIndex.append("==[[" + sTileAll + "]]==\r");


            for (int intSheet = 0; intSheet < workbook.getNumberOfSheets(); intSheet++) {
                HSSFSheet st = workbook.getSheetAt(intSheet);//sheet页面作为二级标题处理
                if (st == null) {
                    continue;
                }

                //计算表格宽度
                float tabWidth = 0;

                if (st.getRow(st.getFirstRowNum()) != null) {
                    for (int intFirstRow = 0; intFirstRow < st.getRow(st.getFirstRowNum()).getLastCellNum(); intFirstRow++) {
                        tabWidth += st.getColumnWidthInPixels(intFirstRow);
                    }
                }

                StringBuilder tmpTable = new StringBuilder();
                tmpTable.append("<table class='wikitable'");
                if (tabWidth > 0) {
                    tmpTable.append(" width='" + tabWidth + "'");
                }
                tmpTable.append(">");

                for (int intRow = 0; intRow < st.getLastRowNum(); intRow++) {
                    HSSFRow row = st.getRow(intRow);
                    if (row == null) {
                        continue;
                    }

                    tmpTable.append("<tr style='height=25'>");

                    for (int intCol = 0; intCol < row.getLastCellNum(); intCol++) {
                        HSSFCell cell = row.getCell(intCol);
                        if (cell == null) {
                            continue;
                        }

                        String tmpNote = readCellToString(cell).isEmpty() ? "&nbsp;" : readCellToString(cell);
                        ;
                        tmpNote = tmpNote.replace("\t", "<BR/>");

                        //列宽度处理
                        String cellWidth = " width='" + st.getColumnWidthInPixels(cell.getColumnIndex()) + "px'";

                        //颜色
                        // style = "color:rgb(255,255,0)
                        String cellColor = "";
                        if (cell.getCellStyle().getFillForegroundColorColor() != null) {
                            cellColor =
                                    " style = \"background-color:"
                                            + cell.getCellStyle().getFillForegroundColorColor().getHexString()
                                            + "\"";
                        }

                        if (cellColor.indexOf("rgb(0,0,0)") >= 0) {
                            cellColor = "";
                        }

                        //合并单元格处理
                        //if (cell. IsMergedCell)
                        if (false) {
                            int rowSpan = 0, colSpan = 0;
                            //readMegCellConfig(st, cell, out rowSpan, out colSpan);

                            if (rowSpan == 1 && colSpan == 1) {
                                //不生成col
                            } else {

                                tmpTable.append("<td");
                                if (colSpan > 1) {
                                    tmpTable.append(" colspan='" + colSpan + "'");
                                }
                                if (rowSpan > 1) {
                                    tmpTable.append(" rowspan='" + rowSpan + "'");
                                }
                                if (cellWidth.isEmpty() == false) {
                                    tmpTable.append(" " + cellWidth);
                                }
                                if (cellColor.isEmpty() == false) {
                                    tmpTable.append(" " + cellColor);
                                }
                                tmpTable.append(">");
                                tmpTable.append(tmpNote);
                                tmpTable.append("</td>");
                            }
                        } else {

                            tmpTable.append("<td ");
                            if (cellWidth.isEmpty() == false) {
                                tmpTable.append(" " + cellWidth);
                            }
                            if (cellColor.isEmpty() == false) {
                                tmpTable.append(" " + cellColor);
                            }
                            tmpTable.append(">");
                            tmpTable.append(tmpNote);
                            tmpTable.append("</td>");
                        }

                    }
                    tmpTable.append("</tr>");

                }
                tmpTable.append("</table>");

                //Regex checkTable = new Regex("<.*?>");
                String strCheck = tmpTable.toString().replaceAll("<.*?>", "");
                strCheck = strCheck.replace("&nbsp;", "");

                if (strCheck.isEmpty() == false) {
                    sOutAll.append("===" + FormatString(st.getSheetName()) + "===\r");
                    sOutAll.append(tmpTable.toString() + "\r");
                }
            }

            op.set_title(sTileAll);
            op.set_text(sOutAll.toString());
            //mOp.Save();
            log.info("写入新页面：" + sTileAll);
        } catch (Exception ex) {
            log.error("Error：" + fileName + "处理失败。需要重新处理。" + ex.getMessage());
        }
    }

    protected String FormatString(String sData) {
        String sR = sData.trim();
        sR = sR.replace("\r\n", "");
        sR = sR.replace("\r", "");
        sR = sR.replace("\t", "");
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

    protected String readCellToString(HSSFCell cell) {
        String strReturn = "";

        CellType ct = cell.getCellType();

        if (ct == CellType.BOOLEAN) {
            strReturn = cell.getBooleanCellValue() ? "TRUE" : "FALSE";
        } else if (ct == CellType.ERROR) {
            strReturn = ErrorEval.getText(cell.getErrorCellValue());
        } else if (ct == CellType.FORMULA) {
            if (cell.getCachedFormulaResultType() == CellType.BOOLEAN) {
                strReturn = cell.getBooleanCellValue() ? "TRUE" : "FALSE";
            }

            if (cell.getCachedFormulaResultType() == CellType.ERROR) {
                strReturn = ErrorEval.getText(cell.getErrorCellValue());
            }

            if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                if (DateUtil.isCellDateFormatted(cell)) {
                    //strReturn = cell.DateCellValue.ToString("yyyy-MM-dd hh:MM:ss");
                    strReturn = cell.toString();
                } else {
                    strReturn = cell.getNumericCellValue() + "";
                }
            }

            if (cell.getCachedFormulaResultType() == CellType.STRING) {
                String str = cell.getStringCellValue();
                if (!str.isEmpty()) {
                    strReturn = str.toString();
                } else {
                    strReturn = null;
                }
            }
        } else if (ct == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                //strReturn = cell.DateCellValue.ToString("yyyy-MM-dd hh:MM:ss");
                strReturn = cell.toString();
            } else {
                strReturn = cell.getNumericCellValue() + "";
            }
        } else if (ct == CellType.STRING) {
            String strValue = cell.getStringCellValue();
            if (strValue.isEmpty()) {
                strReturn = null;
            } else {
                strReturn = strValue.toString();
            }
        }

        return strReturn;
    }

}

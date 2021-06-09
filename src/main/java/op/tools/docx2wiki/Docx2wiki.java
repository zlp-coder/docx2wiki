package op.tools.docx2wiki;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;


public class Docx2wiki {

    public static void main(String args[]) {
        File f = new File("D:\\CompAccountValidate.docx");

        DocTransfer trans = new DocTransfer();
        WikiOperator op = new WikiOperator();

        try {
            //主要的转换方法，输出在 op 对象内
            trans.Transfer(f.getName(), new FileInputStream(f), op);

            //String op.get_text() 转换后的文本内容
            //List<UploadBmpInfo> op.get_bmpInfo() word文档中的图形，可通过saveBmpasFile保存为本地文件

            op.saveBmpasFile("E:\\2806\\");

            File ftxt = new File("E:\\2806\\" + op.get_title() + ".txt");
            FileOutputStream fout = new FileOutputStream(ftxt);
            fout.write(op.get_text().getBytes());
            fout.flush();
            fout.close();

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }


}


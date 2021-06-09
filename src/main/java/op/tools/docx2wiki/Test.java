package op.tools.docx2wiki;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

public class Test {
    public static void main(String[] args) {
        int width = 300;
        int height = 100;
        String text = "爱我中华";
        int x = 0;
        int y = 0;
        BufferedImage processDiagram = new BufferedImage(300, 100,
                BufferedImage.TYPE_INT_ARGB);
        Graphics2D g = (Graphics2D) processDiagram.createGraphics();
        Font font = new Font("宋体", Font.BOLD, 12);
        g.setFont(font);
        FontMetrics fontMetrics = g.getFontMetrics();

        int textX = x + ((width - fontMetrics.stringWidth(text)) / 2);
        int textY = y + ((height - fontMetrics.getHeight()) / 2)
                + fontMetrics.getHeight();

        g.drawString("fuck , 这是中文 这是中文", textX, textY);
        File outFile = new File("E:\\DownLoad\\2806\\统一认证_DDS_V2.12\\newfile.png");
        try {
            ImageIO.write(processDiagram, "png", outFile);
        } catch (IOException e) {
            e.printStackTrace();
        }// 写图片
    }
}

package op.tools.docx2wiki;


import org.freehep.graphicsio.emf.EMFInputStream;
import org.freehep.graphicsio.emf.EMFRenderer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import javax.imageio.stream.ImageOutputStream;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class WikiOperator {
    private static Logger log = LoggerFactory.getLogger(WikiOperator.class);

    private List<UploadBmpInfo> _bmpInfo;
    private String _title;
    private String _text;

    private String _wikiUserName;

    public WikiOperator() {

        this._bmpInfo = new ArrayList<UploadBmpInfo>();
    }

    public void addBmpInfo(UploadBmpInfo bmp) {
        bmp.set_Author(_wikiUserName);
        _bmpInfo.add(bmp);
    }

    //save bmp data to file system. prepared for upload to wiki.
    public void saveBmpasFile(String bmpFileTempPath) {
        if (bmpFileTempPath == null || bmpFileTempPath.isEmpty()) {
            bmpFileTempPath = new PathTools().getPath() + File.separator + "temp";
        }

        try {
            for (UploadBmpInfo info : get_bmpInfo()) {
                File fbmp = new File(bmpFileTempPath);
                if (!fbmp.exists()) {
                    fbmp.mkdirs();
                }

                info.set_FileFullName(bmpFileTempPath + File.separator + info.get_FileName());

                File f = new File(info.get_FileFullName());
                if (f.getName().indexOf(".png") >= 0) {
                    InputStream is = new ByteArrayInputStream(info.get_BmpData());
                    ImageIO.write(ImageIO.read(is), "png", f);
                }

                if (f.getName().indexOf(".jpeg") >= 0) {
                    InputStream is = new ByteArrayInputStream(info.get_BmpData());
                    ImageIO.write(ImageIO.read(is), "jpeg", f);
                }

                if (f.getName().indexOf(".emf") >= 0) {

                    InputStream preis = new ByteArrayInputStream(info.get_BmpData());

                    info.set_FileName(info.get_FileName().replace(".emf", ".png"));
                    info.set_FileFullName(info.get_FileFullName().replace(".emf", ".png"));

                    InputStream is = new ByteArrayInputStream(emfToPng(preis));

                    File ftmp = new File(info.get_FileFullName());
                    ImageIO.write(ImageIO.read(is), "png", ftmp);
                }

                log.info("save bmp file to:" + info.get_FileFullName());
            }
        } catch (Exception ex) {
            log.error(ex.getMessage());
            ex.printStackTrace();
        }
    }
/*
    //这个是从freehep的源码抄过来的,改了入口参数
    //https://github.com/freehep/freehep-vectorgraphics/blob/master/freehep-graphicsio-emf/src/main/java/org/freehep/graphicsio/emf/EMFConverter.java

    private void saveEMFtoPNG(String type, InputStream fin, String destFileName) throws  Exception{
        List<?> exportFileTypes = ExportFileType.getExportFileTypes(type);
        if (exportFileTypes == null || exportFileTypes.size() == 0) {
            System.out.println(
                    type + " library is not available. check your classpath!");
            return;
        }

        ExportFileType exportFileType = (ExportFileType) exportFileTypes.get(0);

        // read the EMF file
//        EMFRenderer emfRenderer = new EMFRenderer(
//                new EMFInputStream(
//                        new FileInputStream(srcFileName)));
        EMFRenderer emfRenderer = new EMFRenderer(
                new EMFInputStream(fin));

        // create the destFileName,
        // replace or add the extension to the destFileName
//        if (destFileName == null || destFileName.length() == 0) {
//            // index of the beginning of the extension
//            int lastPointIndex = srcFileName.lastIndexOf(".");
//
//            // to be sure that the point separates an extension
//            // and is not part of a directory name
//            int lastSeparator1Index = srcFileName.lastIndexOf("/");
//            int lastSeparator2Index = srcFileName.lastIndexOf("\\");
//
//            if (lastSeparator1Index > lastPointIndex ||
//                    lastSeparator2Index > lastPointIndex) {
//                destFileName = srcFileName + ".";
//            } else if (lastPointIndex > -1) {
//                destFileName = srcFileName.substring(
//                        0, lastPointIndex + 1);
//            }
//
//            // add the extension
//            destFileName += type.toLowerCase();
//        }

        // TODO there is no possibility to use Constants of base class!
             create SVG properties
            Properties p = new Properties();
            p.put(SVGGraphics2D.EMBED_FONTS, Boolean.toString(false));
            p.put(SVGGraphics2D.CLIP, Boolean.toString(true));
            p.put(SVGGraphics2D.COMPRESS, Boolean.toString(false));
            p.put(SVGGraphics2D.TEXT_AS_SHAPES, Boolean.toString(false));
            p.put(SVGGraphics2D.FOR, "Freehep EMF2SVG");
            p.put(SVGGraphics2D.TITLE, emfFileName);

        EMFPanel emfPanel = new EMFPanel();
        emfPanel.setRenderer(emfRenderer);

        // TODO why uses this classes components?!
        exportFileType.exportToFile(
                new File(destFileName),
                emfPanel,
                emfPanel,
                null,
                "Freehep EMF converter");
    }
*/

    private byte[] emfToPng(InputStream is) throws Exception {
        byte[] by = null;
        EMFInputStream emf = null;
        EMFRenderer emfRenderer = null;

        ByteArrayOutputStream baos = null;

        ImageOutputStream imageOutputStream = null;
        try {
            emf = new EMFInputStream(is, EMFInputStream.DEFAULT_VERSION);
            emfRenderer = new EMFRenderer(emf);

            //Font ft = new Font("宋体", Font.PLAIN, 12);

            final int width = (int) emf.readHeader().getBounds().getWidth();
            final int height = (int) emf.readHeader().getBounds().getHeight();

            final BufferedImage result = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
            Graphics2D g = (Graphics2D) result.createGraphics();
            //g.setFont(ft);

            emfRenderer.paint(g);
            //emfRenderer.setFont(ft);

            //g.drawString("fuck , 这是中文 这是中文", 100, 100);
            //g.drawString("\u4e2d\u6587",100,120);

            baos = new ByteArrayOutputStream();

            imageOutputStream = ImageIO.createImageOutputStream(baos);

            ImageIO.write(result, "png", imageOutputStream);

            by = baos.toByteArray();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (imageOutputStream != null) {
                    imageOutputStream.close();
                }
                if (baos != null) {
                    baos.close();
                }
                if (emfRenderer != null) {
                    emfRenderer.closeFigure();
                }
                if (emf != null) {
                    emf.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }

        }
        return by;
    }


    public String get_title() {
        return _title;
    }

    public void set_title(String _title) {
        this._title = _title;
    }

    public String get_text() {
        return _text;
    }

    public void set_text(String _text) {
        this._text = _text;
    }

    public void set_wikiUserName(String _wikiUserName) {
        this._wikiUserName = _wikiUserName;
    }

    public List<UploadBmpInfo> get_bmpInfo() {
        return _bmpInfo;
    }

    public void set_bmpInfo(List<UploadBmpInfo> _bmpInfo) {
        this._bmpInfo = _bmpInfo;
    }

}

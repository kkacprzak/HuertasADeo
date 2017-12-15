package fivedots;

/*  Images.java
    Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016
    Heavily edited and formated by K. Kacprzak 2017
    A growing collection of utility functions to make Office easier to use. 
    They are currently divided into the following groups:
    * image I/O
 */

import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameContainer;
import com.sun.star.document.XMimeTypeInfo;
import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;
import com.sun.star.uno.Exception;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.Base64;
import javax.imageio.ImageIO;

public class Images {

    public static String getBitmap(String fnm) {
        try {
            XNameContainer bitmapContainer = Lo.createInstanceMSF(
                    XNameContainer.class, 
                    "com.sun.star.drawing.BitmapTable");
            if (!FileIO.isOpenable(fnm)) return null;
            String picURL = FileIO.fnmToURL(fnm);
            if (picURL == null) return null;
            bitmapContainer.insertByName(fnm, picURL);
            return (String) bitmapContainer.getByName(fnm);
        } catch (Exception e) {
            System.out.println("Could not create a bitmap container for " 
                    + fnm + " " + e);
            return null;
        }
    }  

    public static BufferedImage loadImage(String fnm) {
        BufferedImage image = null;
        try {
            image = ImageIO.read(new File(fnm));
            System.out.println("Loaded " + fnm);
        } catch (java.io.IOException e) {
            System.out.println("Unable to load " + fnm + " " + e);
        }
        return image;
    }  

    public static void saveImage(BufferedImage im, String fnm) {
        if (im == null) {
            System.out.println("No data to save in " + fnm);
            return;
        }
        try {
            ImageIO.write(im, "png", new File(fnm));
            System.out.println("Saved image to file: " + fnm);

        } catch (java.lang.Exception e) {
            System.out.println("Could not save image to " + fnm + ": " + e);
        }
    } 

    public static byte[] im2bytes(BufferedImage im) {
        byte[] bytes = null;
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()){
            ImageIO.write(im, "png", baos);
            baos.flush();
            bytes = baos.toByteArray();
            baos.close();
        } catch (java.io.IOException e) {
            System.out.println("Could not convert image to bytes");
        }
        return bytes;
    }  

    public static String im2String(BufferedImage im) {
        byte[] bytes = im2bytes(im);
        return Base64.getMimeEncoder().encodeToString(bytes);
    }  

    public static BufferedImage string2im(String s) {
        byte[] bytes = Base64.getMimeDecoder().decode(s);
        return bytes2im(bytes);
    }  

    public static BufferedImage bytes2im(byte[] bytes) {
        try {
            ByteArrayInputStream bais = new ByteArrayInputStream(bytes);
            return ImageIO.read(bais);
        } catch (java.io.IOException ioe) {
            System.out.println("Could not convert bytes to image");
            return null;
        }
    }  

    public static XGraphic im2graphic(BufferedImage im) {
        if (im == null) {
            System.out.println("No image found");
            return null;
        }
        String tempFnm = FileIO.createTempFile("png");
        if (tempFnm == null) {
            System.out.println("Could not create a temporary file for the image");
            return null;
        }
        Images.saveImage(im, tempFnm);
        XGraphic graphic = Images.loadGraphicFile(tempFnm);
        FileIO.deleteFile(tempFnm);
        return graphic;
    }  

    public static XGraphic loadGraphicFile(String imFnm) {
        System.out.println("Loading XGraphic from " + imFnm);
        XGraphicProvider gProvider = Lo.createInstanceMCF(XGraphicProvider.class,
                "com.sun.star.graphic.GraphicProvider");
        if (gProvider == null) {
            System.out.println("Graphic Provider could not be found");
            return null;
        }
        PropertyValue[] fileProps = Props.makeProps("URL", FileIO.fnmToURL(imFnm));
        try {
            return gProvider.queryGraphic(fileProps);
        } catch (Exception e) {
            System.out.println("Could not load XGraphic from " 
                    + imFnm+ ": " + e);
            return null;
        }
    }  

    public static Size getSizePixels(String imFnm) {
        XGraphic graphic = loadGraphicFile(imFnm);
        if (graphic == null) return null;
        return (Size) Props.getProperty(graphic, "SizePixel");
    }  

    public static Size getSize100mm(String imFnm) {
        XGraphic graphic = loadGraphicFile(imFnm);
        if (graphic == null) return null;
        return (Size) Props.getProperty(graphic, "Size100thMM");
    }  

    public static XGraphic loadGraphicLink(Object graphicLink) {
        XGraphicProvider gProvider = Lo.createInstanceMCF(
                XGraphicProvider.class,
                "com.sun.star.graphic.GraphicProvider");
        if (gProvider == null) {
            System.out.println("Graphic Provider could not be found");
            return null;
        }
        try {
            XPropertySet xprops = Lo.qi(XPropertySet.class, graphicLink);
            PropertyValue[] gProps = Props.makeProps("URL",
                    (String) xprops.getPropertyValue("GraphicURL"));

            return gProvider.queryGraphic(gProps);
        } catch (Exception e) {
            System.out.println("Unable to retrieve graphic");
            return null;
        }
    } 

    public static BufferedImage graphic2Im(XGraphic graphic) {
        if (graphic == null) {
            System.out.println("No graphic found");
            return null;
        }
        String tempFnm = FileIO.createTempFile("png");
        if (tempFnm == null) {
            System.out.println("Could not create a temporary file for the graphic");
            return null;
        }
        Images.saveGraphic(graphic, tempFnm, "png");
        BufferedImage im = Images.loadImage(tempFnm);
        FileIO.deleteFile(tempFnm);
        return im;
    }  

    public static void saveGraphic(XGraphic pic, String fnm, String imFormat) {
        System.out.println("Saving graphic in " + fnm);
        XGraphicProvider gProvider = Lo.createInstanceMCF(
                XGraphicProvider.class,
                "com.sun.star.graphic.GraphicProvider");
        if (gProvider == null) {
            System.out.println("Graphic Provider could not be found");
            return;
        }
        if (pic == null) {
            System.out.println("Supplied image is null");
            return;
        }
        PropertyValue[] pngProps = Props.makeProps(
                "URL", FileIO.fnmToURL(fnm),
                "MimeType", "image/" + imFormat);
        try {
            gProvider.storeGraphic(pic, pngProps);
        } catch (com.sun.star.uno.Exception e) {
            System.out.println("Unable to save graphic");
        }
    }  

    public static String[] getMimeTypes() {
        XMimeTypeInfo mi = Lo.createInstanceMCF(
                XMimeTypeInfo.class,
                "com.sun.star.drawing.GraphicExportFilter");
        return mi.getSupportedMimeTypeNames();
    }  

    public static String changeToMime(String imFormat) {
        String[] names = getMimeTypes();
        String imf = imFormat.toLowerCase().trim();
        for (String name : names) {
            if (name.contains(imf)) {
                System.out.println("Using mime type: " + name);
                return name;
            }
        }
        System.out.println("No matching mime type, so using image/png");
        return "image/png";
    }  

    public static Size calcScale(String fnm, int maxWidth, int maxHeight) {
        Size imSize = Images.getSize100mm(fnm);   
        if (imSize == null) return null;
        double widthScale = ((double) maxWidth * 100) / imSize.Width;
        double heightScale = ((double) maxHeight * 100) / imSize.Height;
        double scaleFactor = (widthScale < heightScale) ? widthScale : heightScale;
        int w = (int) Math.round(imSize.Width * scaleFactor / 100);
        int h = (int) Math.round(imSize.Height * scaleFactor / 100);
        return new Size(w, h);
    }  

}  


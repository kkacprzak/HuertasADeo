package fivedots;

/*  FileIO.java
    Andrew Davison, ad@fivedots.coe.psu.ac.th, March 2015
    Heavily edited and formated by K. Kacprzak 2017
    A growing collection of utility functions to make Office easier to use.
    They are currently divided into the following groups:
    * File IO
    * file creation / deletion
    * saving/writing to a file
    * zip access
 */

import com.sun.star.container.XNameAccess;
import com.sun.star.io.XActiveDataSink;
import com.sun.star.io.XInputStream;
import com.sun.star.io.XTextInputStream;
import com.sun.star.packages.zip.XZipFileAccess;
import com.sun.star.ucb.XSimpleFileAccess3;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Enumeration;
import java.util.TimeZone;
import java.util.zip.ZipFile;

public class FileIO {

    public static String getUtilsFolder() {
        try {
            return FileIO.class.getProtectionDomain()
                    .getCodeSource()
                    .getLocation()
                    .toURI()
                    .getPath();
        } catch (URISyntaxException e) {
            System.out.println(e);
            return null;
        }
    }

    public static String getAbsolutePath(String fnm) {
        return new File(fnm).getAbsolutePath();
    }

    public static String urlToPath(String url) 
            throws MalformedURLException, URISyntaxException {
        return Paths.get(new URL(url).toURI()).toString();
    }

    public static boolean isOpenable(String fnm) {
        File f = new File(fnm);
        if (!f.exists()) {
            System.out.println(fnm + " does not exist");
            return false;
        }
        if (!f.isFile()) {
            System.out.println(fnm + " is not a file");
            return false;
        }
        if (!f.canRead()) {
            System.out.println(fnm + " is not readable");
            return false;
        }
        return true;
    }

    public static String fnmToURL(String fnm) {
        try {
            StringBuffer sb;
            String path = new File(fnm).getCanonicalPath();
            sb = new StringBuffer("file:///");
            sb.append(path.replace('\\', '/'));
            return sb.toString();
        } catch (java.io.IOException e) {
            System.out.println("Could not access "+fnm+" "+e);
            return null;
        }
    }

    public static String URI2Path(String URIfnm) {
        try {
            File file = new File(new URI(URIfnm).getSchemeSpecificPart());
            return file.getCanonicalPath();
        } catch (URISyntaxException | IOException e) {
            System.out.println("Could not translate settings path "+e);
            return URIfnm;
        }
    }

    public static boolean makeDirectory(String dir) {
        File file = new File(dir);
        if (!file.exists()) {
            if (file.mkdir()) {
                System.out.println("Created " + dir);
                return true;
            } else {
                System.out.println("Could not create " + dir);
                return false;
            }
        } else return true;
    } 

    public static String[] getFileNames(String dir) {
        File[] files = new File(dir).listFiles();
        if (files == null) {
            System.out.println("No directory found called " + dir);
            return null;
        }
        ArrayList<String> results = new ArrayList<>();
        for (File file : files) if (file.isFile()) results.add(file.getName());
        int numFiles = results.size();
        if (numFiles == 0) {
            System.out.println("No files found in the directory " + dir);
            return null;
        }
        String[] fnms = new String[numFiles];
        for (int i = 0; i < numFiles; i++) fnms[i] = results.get(i);
        return fnms;
    } 

    public static String getFnm(String path) {
        return (new File(path)).getName();
    }

    public static String createTempFile(String imFormat) {
        try {
            File temp = File.createTempFile("loTemp", "." + imFormat);
            temp.deleteOnExit();
            return temp.getAbsolutePath();
        } catch (java.io.IOException e) {
            System.out.println("Could not create temp file "+e);
            return null;
        }
    }  

    public static void deleteFiles(ArrayList<String> dbFnms) {
        System.out.println();
        for (int i = 0; i < dbFnms.size(); i++) {
            deleteFile(dbFnms.get(i));
        }
    }

    public static void deleteFile(String fnm) {
        File file = new File(fnm);
        if (file.delete()) System.out.println(fnm + " deleted");
        else System.out.println(fnm + " could not be deleted");
    }

    public static void saveString(String fnm, String s) {
        if (s == null) {
            System.out.println("No data to save in " + fnm);
            return;
        }
        try (FileWriter fw = new FileWriter(new File(fnm))) {
            fw.write(s);
        } catch (java.io.IOException ex) {
            System.out.println("Could not save string to file: " + fnm + " " +ex);
        }
        System.out.println("Saved string to file: " + fnm);
    }

    public static void saveBytes(String fnm, byte[] bytes) {
        if (bytes == null) {
            System.out.println("No data to save in " + fnm);
            return;
        }
        try (FileOutputStream fos = new FileOutputStream(fnm)) {
            fos.write(bytes);
        } catch (java.io.IOException ex) {
            System.out.println("Could not save bytes to file: " + fnm);
        }
        System.out.println("Saved bytes to file: " + fnm);
    }

    public static void saveArray(String fnm, Object[][] arr) {
        if (arr == null) {
            System.out.println("No data to save in " + fnm);
            return;
        }
        try (BufferedWriter bw = new BufferedWriter(new FileWriter(fnm))) {
            int numCols = arr[0].length;
            int numRows = arr.length;
            for (int j = 0; j < numRows; j++) {
                for (int i = 0; i < numCols; i++) 
                    bw.write((String) arr[j][i] + "\t");
                bw.write("\n");
            }
        } catch (java.io.IOException ex) {
            System.out.println("Could not save array to file: " + fnm);
        }
        System.out.println("Saved array to file: " + fnm);
    } 

    public static void saveArray(String fnm, double[][] arr) {
        if (arr == null) {
            System.out.println("No data to save in " + fnm);
            return;
        }
        try (BufferedWriter bw = new BufferedWriter(new FileWriter(fnm))) {
            int numCols = arr[0].length;
            int numRows = arr.length;
            for (int j = 0; j < numRows; j++) {
                for (int i = 0; i < numCols; i++) bw.write(arr[j][i] + "\t");
                bw.write("\n");
            }
        } catch (java.io.IOException ex) {
            System.out.println("Could not save array to file: " + fnm);
        }
        System.out.println("Saved array to file: " + fnm);
    }

    public static void appendTo(String fnm, String msg) {
        try (FileWriter fw = new FileWriter(fnm, true)) {
            fw.write(msg + "\n");
        } catch (java.io.IOException e) {
            System.out.println("Unable to append to " + fnm);
        }
    }

    public static XZipFileAccess zipAccess(String fnm) {
        return Lo.createInstanceMCF(
                XZipFileAccess.class,
                "com.sun.star.packages.zip.ZipFileAccess",
                new Object[]{fnmToURL(fnm)});
    }

    public static void zipListUno(String fnm) {
        XZipFileAccess zfa = zipAccess(fnm);
        XNameAccess nmAccess = Lo.qi(XNameAccess.class, zfa);
        String[] names = nmAccess.getElementNames();
        System.out.println("\nZipped Contents of " + fnm);
        Lo.printNames(names, 1);
    }

    public static void unzipFile(XZipFileAccess zfa, String fnm) {
        String fileName = Info.getName(fnm);
        String ext = Info.getExt(fnm);
        try {
            System.out.println("Extracting " + fnm);
            XInputStream inStream = zfa.getStreamByPattern("*" + fnm);
            XSimpleFileAccess3 fileAcc = Lo.createInstanceMCF(
                    XSimpleFileAccess3.class,
                    "com.sun.star.ucb.SimpleFileAccess");
            String copyFnm = (ext == null) ? (fileName + "Copy")
                    : (fileName + "Copy." + ext);
            System.out.println("Saving to " + copyFnm);
            fileAcc.writeFile(FileIO.fnmToURL(copyFnm), inStream);
        } catch (com.sun.star.uno.Exception e) {
            System.out.println(e);
        }
    }  
    public static String getMimeType(XZipFileAccess zfa) {
        try {
            XInputStream inStream = zfa.getStreamByPattern("mimetype");
            String[] lines = FileIO.readLines(inStream);
            if (lines != null) return lines[0].trim();
        } catch (com.sun.star.uno.Exception e) {
            System.out.println(e);
        }
        System.out.println("No mimetype found");
        return null;
    } 

    public static String[] readLines(XInputStream is) {
        String[] linesArr = null;
        ArrayList<String> lines = new ArrayList<>();
        try {
            XTextInputStream tis = Lo.createInstanceMCF(
                    XTextInputStream.class,
                    "com.sun.star.io.TextInputStream");
            XActiveDataSink sink = Lo.qi(XActiveDataSink.class, tis);
            sink.setInputStream(is);
            while (!tis.isEOF()) lines.add(tis.readLine());
            tis.closeInput();
            linesArr = new String[lines.size()];
            lines.toArray(linesArr);
        } catch (Exception e) {
            System.out.println(e);
        }

        return linesArr;
    }  

    public static void zipList(String fnm) {
        DateFormat df = DateFormat.getDateInstance(); 
        DateFormat tf = DateFormat.getTimeInstance(); 
        tf.setTimeZone(TimeZone.getDefault());
        try {
            ZipFile zfile = new ZipFile(fnm);
            System.out.println("Listing of " + zfile.getName() + ":");
            System.out.println("Raw Size    Size     Date        Time         Name");
            System.out.println("--------  -------  -------      -------      --------");
            Enumeration<? extends java.util.zip.ZipEntry> zfs = zfile.entries();
            while (zfs.hasMoreElements()) {
                java.util.zip.ZipEntry entry = (java.util.zip.ZipEntry) zfs.nextElement();
                Date d = new Date(entry.getTime());
                System.out.print(padSpaces(entry.getSize(), 9) + " ");
                System.out.print(padSpaces(entry.getCompressedSize(), 7) + " ");
                System.out.print(" " + df.format(d) + " ");
                System.out.print(" " + tf.format(d) + "  ");
                System.out.println(" " + entry.getName());
            }
            System.out.println();
        } catch (java.io.IOException e) {
            System.out.println(e);
        }
    } 

    private static String padSpaces(long l, int width) {
        String s = Long.toString(l);
        while (s.length() < width)  s += " ";
        return s;
    }  

} 


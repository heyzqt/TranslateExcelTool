package com.heyzqt;

import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Label;
import jxl.write.biff.RowsExceededException;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;

import java.io.*;
import java.util.Iterator;

public class Main {

    public final static String FILENAME = "E:\\translate.xls";

    //public final static String XMLPATH = "origin_keys/";
    public final static String XMLPATH = "E:\\origin_keys\\";

    private final static String STANDARDNAME = "menu_arrays_";

//    private final static String[] filenames = {"menu_arrays_ar.xml", "menu_arrays_bg_rBG.xml", "menu_arrays_cs.xml",
//            "menu_arrays_da.xml", "menu_arrays_de.xml", "menu_arrays_el_rGR.xml", "menu_arrays_es.xml",
//            "menu_arrays_fa_rIR.xml", "menu_arrays_fi.xml", "menu_arrays_fr.xml", "menu_arrays_hr.xml",
//            "menu_arrays_hu.xml", "menu_arrays_in_rID.xml", "menu_arrays_it.xml", "menu_arrays_iw_rIL.xml",
//            "menu_arrays_mn_rMN.xml", "menu_arrays_ms_rMY.xml", "menu_arrays_my_rMM.xml", "menu_arrays_nl.xml",
//            "menu_arrays_no_rNOR.xml", "menu_arrays_pl.xml", "menu_arrays_pt.xml", "menu_arrays_ro.xml",
//            "menu_arrays_ru.xml", "menu_arrays_sk.xml", "menu_arrays_sl.xml", "menu_arrays_sq_rAL.xml",
//            "menu_arrays_sr.xml", "menu_arrays_sv.xml", "menu_arrays_sw_rTZ.xml", "menu_arrays_ta_rIN.xml",
//            "menu_arrays_th.xml", "menu_arrays_tr.xml", "menu_arrays_uk_rUA.xml", "menu_arrays_vi_rVN.xml"};


    private static boolean isFileExisted = false;

    public static void main(String[] args) {
        // write your code here

        new ToolFrame();

//        String[] countries = {"fr", "pl", "sw_TZ", "zz"};
//        Main.transformEXCEL2XML(FILENAME, Main.XMLPATH, countries,
//                3, 2, 2, 4);
        //removeRow(FILENAME, 1, 35, 4, 90);
        //insertRow(1, 35, 2, 1);
//        addSingleCell(FILENAME, 1, 35,
//                2, 5, 2, 8, 3);
        //copyRowA2RowB(FILENAME, 1, 2, 35, 3, 3);
//        transformEXCEL2XML(FILENAME, XMLPATH, filenames, 1,
//                35, 3, 2, 2, 4);
//        transformEXCEL2XMLArray(FILENAME, XMLPATH, 1, 35
//                , 3, 2, 2, 2);
    }

    public static String[] createFileNames(String filetype, String[] countries) {
        int length = countries.length;
        String[] result = new String[length];
        String filehead;
        switch (filetype) {
            case Constant.FILE_STRINGS:
                filehead = "strings_";
                for (int i = 0; i < length; i++) {
                    result[i] = filehead + countries[i] + ".xml";
                }
                break;
            case Constant.FILE_MENU_STRINGS:
                filehead = "menu_strings_";
                for (int i = 0; i < length; i++) {
                    result[i] = filehead + countries[i] + ".xml";
                }
                break;
            case Constant.FILE_NAV_STRINGS:
                filehead = "nav_strings_";
                for (int i = 0; i < length; i++) {
                    result[i] = filehead + countries[i] + ".xml";
                }
                break;
            case Constant.FILE_TIMESHIFT_STRINGS:
                filehead = "timeshift_strings_";
                for (int i = 0; i < length; i++) {
                    result[i] = filehead + countries[i] + ".xml";
                }
                break;
            case Constant.FILE_MMP_STRINGS:
                filehead = "timeshift_strings_";
                for (int i = 0; i < length; i++) {
                    result[i] = filehead + countries[i] + ".xml";
                }
                break;
            case Constant.FILE_CEC_STRINGS:
                filehead = "cec_strings_";
                for (int i = 0; i < length; i++) {
                    result[i] = filehead + countries[i] + ".xml";
                }
                break;
            case Constant.FILE_THR_MENU_STRINGS:
                filehead = "thr_menu_strings_";
                for (int i = 0; i < length; i++) {
                    result[i] = filehead + countries[i] + ".xml";
                }
                break;
        }
        return result;
    }

    public static String createFileNames(String[] filenames, String regex, String replacement) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < filenames.length; i++) {
            sb.append("\"");
            String temp = filenames[i].replaceAll(regex, replacement);
            sb.append(temp);
            sb.append("\",");
        }
        System.out.println(sb);
        return sb.toString();
    }

    /**
     * 将Excel转换为Xml文件(array形式)
     *
     * @param excelPath        Excel文件路径，如 "E:\nav_strings_keys.xls"
     * @param xmlPath          XML文件写入路径，如 "origin_keys/"
     * @param beginSheetIndex  开始表序号 (表序号从1开始)
     * @param endSheetIndex    结束表序号 (表序号从1开始)
     * @param keyColumnIndex   key值列序号（列序号从1开始）
     * @param valueColumnIndex value值列序号（列序号从1开始）
     * @param beginRowIndex    excel开始写入的行序号(行序号从1开始),beginRowIndex会被写入
     * @param endRowIndex      excel结束写入的行序号(行序号从1开始),bendRowIndex会被写入
     * @return
     */
    public static boolean transformEXCEL2XMLArray(String excelPath, String xmlPath, int beginSheetIndex, int
            endSheetIndex, int keyColumnIndex, int valueColumnIndex, int beginRowIndex, int endRowIndex) {
        beginSheetIndex -= 1;
        endSheetIndex -= 1;
        beginRowIndex -= 1;
        endRowIndex -= 1;
        keyColumnIndex -= 1;
        valueColumnIndex -= 1;
        File file = new File(excelPath);
        String[] keys = new String[endRowIndex - beginRowIndex + 1];
        String[][] keys_values = new String[endRowIndex - beginRowIndex + 1][endSheetIndex - beginSheetIndex + 1];

        try {
            System.out.println("transform begin");
            ToolFrame.showLog("transform begin");
            InputStream in = new FileInputStream(file);
            WorkbookSettings settings = new WorkbookSettings();
            //保证读取（read）excel的编码格式和写入（write）的编码格式统一，避免乱码
            settings.setEncoding("ISO-8859-1");
            Workbook workbook = Workbook.getWorkbook(in, settings);

            Sheet[] sheets = workbook.getSheets();
            int x = 0;
            //find keys
            for (int i = beginRowIndex; i <= endRowIndex; i++) {
                Sheet sheet = workbook.getSheet(0);
                keys[x] = sheet.getCell(keyColumnIndex, i).getContents();
                System.out.println(keys[x]);
                x++;
            }

            System.out.println();

            //find values
            int row = 0;
            int col = 0;
            int index = 0;
            int m = 0;
            int n = 0;
            for (int i = beginSheetIndex; i <= endSheetIndex; i++) {
                System.out.println("i = " + i + " sheet name = " + sheets[i].getName());
                for (int j = beginRowIndex; j <= endRowIndex; j++) {
//                    System.out.println("row " + (j + 1) + " : " + sheets[i].getCell(keyColumnIndex, j)
//                            .getContents() +
//                            "," + sheets[i].getCell(valueColumnIndex, j).getContents());


                    if (j == beginRowIndex) {
                        keys_values[0][index] = sheets[i].getCell(valueColumnIndex, j).getContents();
                    } else if (j == beginRowIndex + 1) {
                        keys_values[1][index] = sheets[i].getCell(valueColumnIndex, j).getContents();
                    } else if (j == beginRowIndex + 2) {
                        keys_values[2][index] = sheets[i].getCell(valueColumnIndex, j).getContents();
                    } else if (j == beginRowIndex + 3) {
                        keys_values[3][index] = sheets[i].getCell(valueColumnIndex, j).getContents();
                    } else if (j == beginRowIndex + 4) {
                        keys_values[4][index] = sheets[i].getCell(valueColumnIndex, j).getContents();
                    }
                }
                index++;
                System.out.println();
                ToolFrame.showLog("");
            }

            System.out.println("transform end");
            ToolFrame.showLog("transform end");

            System.out.println("二维数组：");
            for (int i = 0; i < keys_values.length; i++) {
                for (int j = 0; j < keys_values[0].length; j++) {
                    System.out.println(j + "  " + keys_values[i][j] + ",");
                }
                System.out.println();
            }


            //change excel to xml
            int length = endSheetIndex - beginSheetIndex + 1;
            for (int i = 0; i < length; i++) {
                String[] values = new String[5];
                values[0] = keys_values[0][i];
                values[1] = keys_values[1][i];
                values[2] = keys_values[2][i];
                values[3] = keys_values[3][i];
                values[4] = keys_values[4][i];
                //write2XMLArray(xmlPath, filenames[i], beginRowIndex, endRowIndex, keys, values);
            }
            return true;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            ToolFrame.showLog("error！！！文件正在被占用，请先关闭文件。");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }
        return false;
    }

    /**
     * 将Excel转换为Xml文件
     *
     * @param excelPath        Excel文件路径，如 "E:\nav_strings_keys.xls"
     * @param xmlPath          XML文件写入路径，如 "origin_keys/"
     * @param beginSheetIndex  开始表序号 (表序号从1开始)
     * @param endSheetIndex    结束表序号 (表序号从1开始)
     * @param keyColumnIndex   key值列序号（列序号从1开始）
     * @param valueColumnIndex value值列序号（列序号从1开始）
     * @param beginRowIndex    excel开始写入的行序号(行序号从1开始),beginRowIndex会被写入
     * @param endRowIndex      excel结束写入的行序号(行序号从1开始),bendRowIndex会被写入
     * @return
     */
    public static boolean transformEXCEL2XML(String excelPath, String xmlPath, String[] filenames, int
            beginSheetIndex, int endSheetIndex, int keyColumnIndex, int valueColumnIndex,
                                             int beginRowIndex, int endRowIndex) {
        beginSheetIndex -= 1;
        endSheetIndex -= 1;
        beginRowIndex -= 1;
        endRowIndex -= 1;
        keyColumnIndex -= 1;
        valueColumnIndex -= 1;
        File file = new File(excelPath);
        String[] keys = new String[endRowIndex - beginRowIndex + 1];
        String[][] keys_values = new String[endRowIndex - beginRowIndex + 1][endSheetIndex - beginSheetIndex + 1];

        try {
            System.out.println("transform begin");
            ToolFrame.showLog("transform begin");
            InputStream in = new FileInputStream(file);
            WorkbookSettings settings = new WorkbookSettings();
            //保证读取（read）excel的编码格式和写入（write）的编码格式统一，避免乱码
            settings.setEncoding("ISO-8859-1");
            Workbook workbook = Workbook.getWorkbook(in, settings);

            if (beginSheetIndex < 0) {
                System.out.println("error beginSheetIndex参数有误");
                ToolFrame.showLog("error beginSheetIndex参数有误，请重新输入！");
                return false;
            }
            if (endSheetIndex >= workbook.getNumberOfSheets()) {
                System.out.println("error endSheetIndex参数有误");
                ToolFrame.showLog("error endSheetIndex参数有误，请重新输入！");
                return false;
            }
            if (beginSheetIndex > endSheetIndex) {
                System.out.println("error sheetIndex参数有误");
                ToolFrame.showLog("error sheetIndex参数有误，请重新输入！");
                return false;
            }

            Sheet[] sheets = workbook.getSheets();
            int x = 0;
            //find keys
            for (int i = beginRowIndex; i <= endRowIndex; i++) {
                Sheet sheet = workbook.getSheet(0);
                keys[x] = sheet.getCell(keyColumnIndex, i).getContents();
                System.out.println(keys[x]);
                x++;
            }

            System.out.println();

            //find values
            int index = 0;
            for (int i = beginSheetIndex; i <= endSheetIndex; i++) {
                //System.out.println("i = " + i + " sheet name = " + sheets[i].getName());
                //ToolFrame.showLog("i = " + i + " sheet name = " + sheets[i].getName());
                int row_index = 0;
                for (int j = beginRowIndex; j <= endRowIndex; j++) {
//                    System.out.println("row " + (j + 1) + " : " + sheets[i].getCell(keyColumnIndex, j)
//                            .getContents() +
//                            "," + sheets[i].getCell(valueColumnIndex, j).getContents());

                    keys_values[row_index++][index] = sheets[i].getCell(valueColumnIndex, j).getContents();

//                    //这里根据需要添加的行数添加,这里我们添加2行
//                    if (j == beginRowIndex) {
//                        keys_values[0][index] = sheets[i].getCell(valueColumnIndex, j).getContents();
//                    } else if (j == beginRowIndex + 1) {
//                        keys_values[1][index] = sheets[i].getCell(valueColumnIndex, j).getContents();
//                    }
                }
                index++;
            }

            System.out.println();
            System.out.println("transform end");
            ToolFrame.showLog("transform end");

            System.out.println("二维数组：");
            ToolFrame.showLog("二维数组：");
            for (int i = 0; i < keys_values.length; i++) {
                for (int j = 0; j < keys_values[0].length; j++) {
                    System.out.print(keys_values[i][j] + ",");
                    ToolFrame.showLogInfo(keys_values[i][j] + ",");
                }
                System.out.println();
                ToolFrame.showLog("");
            }


            //change excel to xml
            int length = endSheetIndex - beginSheetIndex + 1;
            int lines = endRowIndex - beginRowIndex + 1;
            System.out.println("要修改的行数 lines = " + lines);
            ToolFrame.showLog("要修改的行数 lines = " + lines);
            for (int i = 0; i < length; i++) {
                String[] values = new String[lines];
                for (int j = 0; j < lines; j++) {
                    values[j] = keys_values[j][i];
                }
//                //这里根据需要添加的行数添加,这里我们添加2行
//                String[] values = new String[2];
//                values[0] = keys_values[0][i];
//                values[1] = keys_values[1][i];
                if (!isFileExisted) {
                    write2XML(xmlPath, filenames[i], beginRowIndex, endRowIndex, keys, values);
                }
            }
            return true;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            ToolFrame.showLog("error！！！文件正在被占用，请先关闭文件。");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }
        return false;
    }

    public static boolean transformEXCEL2XML(String excelPath, String xmlPath, String[] filenames,
                                             String[] countriesname,
                                             int keyColumnIndex, int valueColumnIndex,
                                             int beginRowIndex, int endRowIndex) {
        beginRowIndex -= 1;
        endRowIndex -= 1;
        keyColumnIndex -= 1;
        valueColumnIndex -= 1;
        File file = new File(excelPath);
        String[] keys = new String[endRowIndex - beginRowIndex + 1];
        int countriesLength = countriesname.length;
        int[] sheets_index = new int[countriesLength];
        for (int i = 0; i < sheets_index.length; i++) {
            sheets_index[i] = -1;
        }
        //first index : the number of keys
        //second index : the number of values
        String[][] keys_values = new String[endRowIndex - beginRowIndex + 1][countriesLength];

        try {
            System.out.println("transform begin");
            ToolFrame.showLog("transform begin");
            InputStream in = new FileInputStream(file);
            WorkbookSettings settings = new WorkbookSettings();
            //保证读取（read）excel的编码格式和写入（write）的编码格式统一，避免乱码
            settings.setEncoding("ISO-8859-1");
            Workbook workbook = Workbook.getWorkbook(in, settings);

            Sheet[] sheets = workbook.getSheets();
            //find the sheets index which are needed to transform to XML file
            int m = 0;
            int flag = 0;
            int[] recordCountriesIndex = new int[countriesLength];
            for (int i = 0; i < recordCountriesIndex.length; i++) {
                recordCountriesIndex[i] = -1;
            }
            for (int i = 0; i < countriesname.length; i++) {
                for (int j = flag; j < sheets.length; j++) {
                    int check = sheets[j].getName().indexOf(countriesname[i]);
                    if (check != -1) {
                        sheets_index[m++] = j;
                        recordCountriesIndex[i] = 1;
                        flag = j;
                        flag++;
                        break;
                    }
                }
            }


            //print selected sheets index
            System.out.println("selected index : ");
            ToolFrame.showLog("selected index : ");
            boolean ifExited = false;
            int n = 0;
            for (int i = 0; i < sheets_index.length; i++) {
                System.out.print("  " + sheets_index[i]);
                ToolFrame.showLogInfo("  " + sheets_index[i]);
            }
            System.out.println();
            for (int i = 0; i < recordCountriesIndex.length; i++) {
                if (recordCountriesIndex[i] == -1) {
                    System.out.println("error!!! There is no such country in the Excel : " + countriesname[i]);
                    ToolFrame.showLog("error!!! There is no such country in the Excel : " + countriesname[i]);
                    ifExited = true;
                }
            }
            if (ifExited) {
                return false;
            }
            System.out.println();

            //find keys
            int x = 0;
            for (int i = beginRowIndex; i <= endRowIndex; i++) {
                Sheet sheet = workbook.getSheet(0);
                keys[x] = sheet.getCell(keyColumnIndex, i).getContents();
                System.out.println(keys[x]);
                x++;
            }

            System.out.println();

            //find values
            int index = 0;
            for (int i = 0; i < countriesLength; i++) {
                System.out.println("i = " + i + " sheet name = " + sheets[sheets_index[i]].getName());
                ToolFrame.showLog("i = " + i + " sheet name = " + sheets[sheets_index[i]].getName());
                int row_index = 0;
                for (int j = beginRowIndex; j <= endRowIndex; j++) {
//                    System.out.println("row " + (j + 1) + " : " + sheets[i].getCell(keyColumnIndex, j)
//                            .getContents() +
//                            "," + sheets[i].getCell(valueColumnIndex, j).getContents());

                    keys_values[row_index++][index] = sheets[sheets_index[i]]
                            .getCell(valueColumnIndex, j).getContents();

//                    //这里根据需要添加的行数添加,这里我们添加2行
//                    if (j == beginRowIndex) {
//                        keys_values[0][index] = sheets[i].getCell(valueColumnIndex, j).getContents();
//                    } else if (j == beginRowIndex + 1) {
//                        keys_values[1][index] = sheets[i].getCell(valueColumnIndex, j).getContents();
//                    }
                }
                index++;
            }

            System.out.println("二维数组：");
            ToolFrame.showLog("二维数组：\n");
            for (int i = 0; i < keys_values.length; i++) {
                for (int j = 0; j < keys_values[0].length; j++) {
                    System.out.print(keys_values[i][j] + ", ");
                    ToolFrame.showLogInfo(keys_values[i][j] + ",");
                }
                System.out.println();
                ToolFrame.showLog("");
            }

            //change excel to xml
            int lines = endRowIndex - beginRowIndex + 1;
            System.out.println("要修改的行数 lines = " + lines);
            for (int i = 0; i < countriesLength; i++) {
                String[] values = new String[lines];
                for (int j = 0; j < lines; j++) {
                    values[j] = keys_values[j][i];
                }
//                //这里根据需要添加的行数添加,这里我们添加2行
//                String[] values = new String[2];
//                values[0] = keys_values[0][i];
//                values[1] = keys_values[1][i];
                if (!isFileExisted) {
                    write2XML(xmlPath, filenames[i], beginRowIndex, endRowIndex, keys, values);
                }
            }
            System.out.println();
            System.out.println("transform end");
            ToolFrame.showLog("transform end");
            return true;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            ToolFrame.showLog("error！！！文件正在被占用，请先关闭文件。");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }
        return false;
    }

    private static void write2XML(String xmlpath, String filename, int beginRowColumn, int endRowColumn, String[]
            keys, String[] values) {
//        beginRowColumn -= 1;
//        endRowColumn -= 1;

        //create xml file
        File newFile = new File(xmlpath + filename);
        if (!newFile.exists()) {
            try {
                newFile.createNewFile();
                //create root element
                PrintStream ps = new PrintStream(new FileOutputStream(newFile));
                ps.println("<resources>\n</resources>");// 往文件里写入字符串
                ps.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            System.out.println("error!!! please remove all the xml files!");
            ToolFrame.showLog("error!!! please remove all the xml files!");
            isFileExisted = true;
            return;
        }

        //添加结点
        SAXReader reader = new SAXReader();
        // 通过read方法读取一个文件 转换成Document对象
        Document document = null;
        try {
            document = reader.read(newFile);
            Element root = document.getRootElement();
            int length = endRowColumn - beginRowColumn + 1;

            for (int i = 0; i < length; i++) {
                root.addElement("string");
            }

            Iterator iterator = root.elementIterator();
            int j = 0;
            while (iterator.hasNext()) {
                Element element = (Element) iterator.next();
                element.addAttribute("name", keys[j]);
                element.setText(values[j]);
                j++;
            }

        } catch (DocumentException e) {
            e.printStackTrace();
            System.out.println("error!!! please remove all the xml files!");
        } catch (ArrayIndexOutOfBoundsException e1) {
            e1.printStackTrace();
            System.out.println("error!!! please remove all the xml files!");
        }

        try {
            writer(document, xmlpath + filename);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void write2XMLArray(String xmlpath, String filename, int beginRowColumn, int endRowColumn, String[]
            keys, String[] values) {
//        beginRowColumn -= 1;
//        endRowColumn -= 1;

        //create xml file
        File newFile = new File(xmlpath + filename);
        if (!newFile.exists()) {
            try {
                newFile.createNewFile();
                //create root element
                PrintStream ps = new PrintStream(new FileOutputStream(newFile));
                ps.println("<resources>\n</resources>");// 往文件里写入字符串
                ps.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        //添加结点
        SAXReader reader = new SAXReader();
        // 通过read方法读取一个文件 转换成Document对象
        Document document = null;
        try {
            document = reader.read(newFile);
            Element root = document.getRootElement();
            int length = endRowColumn - beginRowColumn + 1;

            for (int i = 0; i < length; i++) {
                root.addElement("string-array");
            }

            Iterator iterator = root.elementIterator();
            int j = 0;
            while (iterator.hasNext()) {
                Element element = (Element) iterator.next();
                element.addAttribute("name", keys[j]);
                Element subElement = element.addElement("item");
                subElement.setText(values[j]);
                j++;
            }

        } catch (DocumentException e) {
            e.printStackTrace();
        }

        try {
            writer(document, xmlpath + filename);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 复制单元格方法
     *
     * @param beginSheetIndex 开始表序号(表序号从1开始)
     * @param endSheetIndex   结束表序号(表序号从1开始)
     * @param readcolumn      读取的单元格列数（序号从1开始）
     * @param readrow         读取的单元格行数（序号从1开始）
     * @param begincolumn     开始写入的单元格列数（序号从1开始）
     * @param beginrow        开始写入的单元格行数（序号从1开始）
     * @param lines           写入多少行
     */
    public static void addSingleCell(String filepath, int beginSheetIndex, int endSheetIndex, int readcolumn,
                                     int readrow, int begincolumn, int beginrow, int lines) {
        beginSheetIndex -= 1;
        endSheetIndex -= 1;
        readcolumn -= 1;
        readrow -= 1;
        begincolumn -= 1;
        beginrow -= 1;
        File file = new File(filepath);
        try {
            System.out.println("copy cell begin");
            ToolFrame.showLog("copy cell begin");
            InputStream in = new FileInputStream(file);
            WorkbookSettings settings = new WorkbookSettings();
            //保证读取（read）excel的编码格式和写入（write）的编码格式统一，避免乱码
            settings.setEncoding("ISO-8859-1");
            //创建工作簿
            Workbook workbook = Workbook.getWorkbook(in, settings);

            if (beginSheetIndex < 0) {
                System.out.println("error beginSheetIndex参数有误");
                ToolFrame.showLog("error beginSheetIndex参数有误，请重新输入！");
                return;
            }
            if (endSheetIndex >= workbook.getNumberOfSheets()) {
                System.out.println("error endSheetIndex参数有误");
                ToolFrame.showLog("error endSheetIndex参数有误，请重新输入！");
                return;
            }
            if (beginSheetIndex > endSheetIndex) {
                System.out.println("error sheetIndex参数有误");
                ToolFrame.showLog("error sheetIndex参数有误，请重新输入！");
                return;
            }

            //创建可写入的工作簿,根据book创建一个操作对象
            WritableWorkbook writableWorkbook = Workbook.createWorkbook(file, workbook, settings);

            for (int i = beginSheetIndex; i <= endSheetIndex; i++) {
                WritableSheet sheet = writableWorkbook.getSheet(i);
                String temp = sheet.getCell(readcolumn, readrow).getContents();
                System.out.println("getCell = " + temp);
                ToolFrame.showLog("getCell = " + temp);
                for (int j = 0; j < lines; j++) {
                    Label label = new Label(begincolumn, beginrow + j, temp);
                    sheet.addCell(label);
                    System.out.println("setCell = " + label.getContents());
                    ToolFrame.showLog("i = " + i + "setCell = " + label.getContents());
                }
                System.out.println();
                ToolFrame.showLog("");
            }

            writableWorkbook.write();
            writableWorkbook.close();
            System.out.println("copy cell end");
            ToolFrame.showLog("copy cell end");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            ToolFrame.showLog("error！！！文件正在被占用，请先关闭文件。");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }
    }

    /**
     * 插入连续行数（空行）
     *
     * @param beginSheetIndex 开始表序号 (从1开始)
     * @param endSheetIndex   结束表序号 (从1开始)
     * @param beginRow        从哪行开始插入(从1开始),在beginRow+1行插入lines行
     * @param lines           插入几行
     */
    public static void insertRow(String filepath, int beginSheetIndex, int endSheetIndex, int beginRow, int lines) {
        beginSheetIndex -= 1;
        endSheetIndex -= 1;
        File file = new File(filepath);
        try {
            System.out.println("insert begin");
            ToolFrame.showLog("insert begin");
            InputStream in = new FileInputStream(file);
            WorkbookSettings settings = new WorkbookSettings();
            //保证读取（read）excel的编码格式和写入（write）的编码格式统一，避免乱码
            settings.setEncoding("ISO-8859-1");
            //创建工作簿
            Workbook workbook = Workbook.getWorkbook(in, settings);

            if (beginSheetIndex < 0) {
                System.out.println("error beginSheetIndex参数有误");
                ToolFrame.showLog("error beginSheetIndex参数有误，请重新输入！");
                return;
            }
            if (endSheetIndex >= workbook.getNumberOfSheets()) {
                System.out.println("error endSheetIndex参数有误");
                ToolFrame.showLog("error endSheetIndex参数有误，请重新输入！");
                return;
            }
            if (beginSheetIndex > endSheetIndex) {
                System.out.println("error sheetIndex参数有误");
                ToolFrame.showLog("error sheetIndex参数有误，请重新输入！");
                return;
            }
            if (lines <= 0) {
                System.out.println("error 插入行数参数有误");
                ToolFrame.showLog("error 插入行数参数有误!");
                return;
            }

            //创建可写入的工作簿,根据book创建一个操作对象
            WritableWorkbook writableWorkbook = Workbook.createWorkbook(file, workbook, settings);

            int sheetsLength = endSheetIndex - beginSheetIndex + 1;    //修改表数目
            for (int i = beginSheetIndex; i < sheetsLength; i++) {
                WritableSheet writableSheet = writableWorkbook.getSheet(i);
                for (int j = 0; j < lines; j++) {
                    writableSheet.insertRow(beginRow);
                }
            }

            writableWorkbook.write();
            writableWorkbook.close();
            System.out.println("insert end");
            ToolFrame.showLog("insert end");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            ToolFrame.showLog("error！！！文件正在被占用，请先关闭文件。");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }
    }

    /**
     * * 删除连续整行
     *
     * @param filepath        excel文件路径
     * @param beginSheetIndex 开始表序号 (表序号从1开始)
     * @param endSheetIndex   结束表序号(表序号从1开始)
     * @param beginRow        从哪行开始删除(行序号从1开始),beginRow会被删除
     * @param endRow          从哪行结束(行序号从1开始),endRow会被删除
     */
    public static void removeRow(String filepath, int beginSheetIndex, int endSheetIndex, int beginRow, int endRow) {
        beginSheetIndex -= 1;
        endSheetIndex -= 1;
        beginRow -= 1;
        int lines = endRow - beginRow;

        File file = new File(filepath);
        try {
            System.out.println("remove begin");
            ToolFrame.showLog("remove begin");
            InputStream in = new FileInputStream(file);
            WorkbookSettings settings = new WorkbookSettings();
            //保证读取（read）excel的编码格式和写入（write）的编码格式统一，避免乱码
            settings.setEncoding("ISO-8859-1");
            //创建工作簿
            Workbook workbook = Workbook.getWorkbook(in, settings);

            if (beginSheetIndex < 0) {
                System.out.println("error beginSheetIndex参数有误");
                ToolFrame.showLog("error beginSheetIndex参数有误，请重新输入！");
                return;
            }
            if (endSheetIndex >= workbook.getNumberOfSheets()) {
                System.out.println("error endSheetIndex参数有误");
                ToolFrame.showLog("error endSheetIndex参数有误，请重新输入！");
                return;
            }
            if (beginSheetIndex > endSheetIndex) {
                System.out.println("error sheetIndex参数有误");
                ToolFrame.showLog("error sheetIndex参数有误，请重新输入！");
                return;
            }

            //创建可写入的工作簿,根据book创建一个操作对象
            WritableWorkbook writableWorkbook = Workbook.createWorkbook(file, workbook, settings);

            for (int i = beginSheetIndex; i <= endSheetIndex; i++) {
                WritableSheet writableSheet = writableWorkbook.getSheet(i);
                for (int j = 0; j < lines; j++) {
                    writableSheet.removeRow(beginRow);
                }
            }

            writableWorkbook.write();
            writableWorkbook.close();
            System.out.println("remove end");
            ToolFrame.showLog("remove end");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            ToolFrame.showLog("error！！！文件正在被占用，请先关闭文件。");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }
    }


    /**
     * 整列复制
     * 将表readSheetIndex的readColume列数据，复制到beginSheetIndex表到endSheetIndex的writeColume列中
     *
     * @param readSheetIndex  读取哪张表(表序号从1开始)
     * @param beginSheetIndex 从哪张表开始写入(表序号从1开始)
     * @param endSheetIndex   写入结束的表的序号(表序号从1开始)
     * @param readColume      读取列序号(行序号从1开始)
     * @param writeColume     写入列序号(行序号从1开始)
     */
    public static void copyRowA2RowB(String filepath, int readSheetIndex, int beginSheetIndex, int endSheetIndex, int
            readColume, int writeColume) {
        readSheetIndex -= 1;
        beginSheetIndex -= 1;
        endSheetIndex -= 1;
        readColume -= 1;
        writeColume -= 1;
        File file = new File(filepath);
        try {
            System.out.println("copy begin");
            ToolFrame.showLog("copy begin");
            InputStream in = new FileInputStream(file);
            WorkbookSettings settings = new WorkbookSettings();
            //保证读取（read）excel的编码格式和写入（write）的编码格式统一，避免乱码
            settings.setEncoding("ISO-8859-1");
            //创建工作簿
            Workbook workbook = Workbook.getWorkbook(in, settings);

            if (beginSheetIndex < 0) {
                System.out.println("error beginSheetIndex参数有误");
                ToolFrame.showLog("error beginSheetIndex参数有误，请重新输入！");
                return;
            }
            if (endSheetIndex >= workbook.getNumberOfSheets()) {
                System.out.println("error endSheetIndex参数有误");
                ToolFrame.showLog("error endSheetIndex参数有误，请重新输入！");
                return;
            }
            if (beginSheetIndex > endSheetIndex) {
                System.out.println("error sheetIndex参数有误");
                ToolFrame.showLog("error sheetIndex参数有误，请重新输入！");
                return;
            }

            //创建可写入的工作簿,根据book创建一个操作对象
            WritableWorkbook writableWorkbook = Workbook.createWorkbook(file, workbook, settings);

            Sheet readSheet = writableWorkbook.getSheet(readSheetIndex);
            //int sheetsLength = endSheetIndex - beginSheetIndex + 1;    //修改表数目
            //单元格内容过长自动换行
            WritableCellFormat cellFormat = new WritableCellFormat();
            cellFormat.setWrap(true);
            for (int i = beginSheetIndex; i <= endSheetIndex; i++) {
                WritableSheet writableSheet = writableWorkbook.getSheet(i);
                //设置列宽
                writableSheet.setColumnView(writeColume, 30);
                for (int j = 0; j < readSheet.getRows(); j++) {
                    Label label = new Label(writeColume, j, readSheet.getCell(readColume, j).getContents(), cellFormat);
                    writableSheet.addCell(label);
                }
            }

            writableWorkbook.write();
            writableWorkbook.close();
            System.out.println("copy end");
            ToolFrame.showLog("copy end");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            ToolFrame.showLog("error！！！文件正在被占用，请先关闭文件。");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }
    }

    /**
     * 把document对象写入新的文件
     *
     * @param document
     * @throws Exception
     */
    public static void writer(Document document, String filepath) throws Exception {
        // 紧凑的格式
        // OutputFormat format = OutputFormat.createCompactFormat();
        // 排版缩进的格式
        OutputFormat format = OutputFormat.createPrettyPrint();
        // 设置编码
        //format.setEncoding("UTF-8");
        // 创建XMLWriter对象,指定了写出文件及编码格式
        // XMLWriter writer = new XMLWriter(new FileWriter(new
        // File("src//a.xml")),format);
//        XMLWriter writer = new XMLWriter(new OutputStreamWriter(
//                new FileOutputStream(new File("src//a.xml")), "UTF-8"), format);
        XMLWriter xmlWriter = new XMLWriter(new OutputStreamWriter(new FileOutputStream(new File(filepath))),
                format);
        // 写入
        xmlWriter.write(document);
        // 立即写入
        xmlWriter.flush();
        // 关闭操作
        xmlWriter.close();
    }
}
package net.inull.ExcelRead;

import java.io.BufferedOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Hashtable;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.EncodeHintType;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.WriterException;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.qrcode.decoder.ErrorCorrectionLevel;

public class Main
{

  public static void main(String[] args)
  {
    try
    {
      // 读取 Excel 文件
      ReadXLSXFile();
    }
    catch (IOException e)
    {
      e.printStackTrace();
    }

  }

  public static void ReadXLSXFile() throws IOException
  {
    // 流输入
    InputStream inputExcelFile = new FileInputStream("D:\\workspaces\\java\\console\\ExcelRead\\file.xlsx");

    // HSSFWorkbook workbook = new HSSFWorkbook(ExcelFileToRead);
    XSSFWorkbook workbook = new XSSFWorkbook(inputExcelFile);

    // HSSFSheet sheet = workbook.getSheetAt(0);
    XSSFSheet sheet = workbook.getSheetAt(0);

//    demo 
//    Iterator<Row> it = sheet.iterator();
//    while (it.hasNext())
//    {
//      Row row = it.next();
//      Cell cell = row.getCell(1);
//      String strFilename = cell.getStringCellValue();
//      System.out.println(strFilename);
//    }

    // 遍历指定 sheet 的所有 row
    for (Row row : sheet)
    {
      // 从第2行开始
      if (row.getRowNum() >= 2)
      {
        Cell cell = row.getCell(1);
        String strFilename = cell.getStringCellValue();

        cell = row.getCell(2);
        String text = cell.getStringCellValue();

        System.out.println(String.valueOf(row.getRowNum()) + "\t" + strFilename + "\t" + text);

        int width = 300;
        int height = 300;

        String strPath = "D:\\workspaces\\java\\console\\ExcelRead\\out\\";
        // 写出二维码
        WriteRCode(strPath + strFilename + ".png", BarcodeFormat.QR_CODE, text, width, height);
      }

//      for (Cell cell : row)
//      {
//        switch (cell.getCellType())
//        {
//          case CellType.NUMERIC:
//          {
//            double numericCellValue = cell.getNumericCellValue();
//            System.out.print(numericCellValue + "\t");
//            break;
//          }
//          case CellType.STRING:
//          {
//            String str = cell.getStringCellValue();
//            System.out.print(str + "\t");
//            break;
//          }
//
//          default:
//            break;
//        }
//      }

      System.out.println("");
    }
  }

  public static void WriteRCode(String fileName, BarcodeFormat typeCode, String strContents, int width, int height)
  {
    Hashtable<EncodeHintType, Object> hints = new Hashtable<EncodeHintType, Object>();
    // 字符集
    hints.put(EncodeHintType.CHARACTER_SET, "GBK");// Constant.CHARACTER);
    // 容错质量
    hints.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.M);

    try
    {
      // 尺寸
      BitMatrix bitMatrix = new MultiFormatWriter().encode(strContents, typeCode, width, height, hints);
      BufferedOutputStream buffer = new BufferedOutputStream(new FileOutputStream(fileName));
      // ok
      MatrixToImageWriter.writeToStream(bitMatrix, "png", new FileOutputStream(new String(fileName)));
      buffer.close();
    }
    catch (WriterException e)
    {
      e.printStackTrace();
    }
    catch (IOException e)
    {
      e.printStackTrace();
    }

  }

}

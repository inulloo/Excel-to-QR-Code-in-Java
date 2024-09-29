package net.inull.ExcelRead;

import java.io.IOException;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExcelRead1Application
{

  public static void main(String[] args)
  {

    // System.out.print("FFF");
    // SpringApplication.run(ExcelRead1Application.class, args);

    try
    {
      Main.ReadXLSXFile();
    }
    catch (IOException e)
    {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
  }

}

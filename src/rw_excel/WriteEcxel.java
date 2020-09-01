package rw_excel;

// *****  IMPORTACIONES ****
import java.io.File;
import java.io.FileOutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.IndexedColors; 
import org.apache.poi.ss.util.CellRangeAddress;

public class WriteEcxel {
    
    public XSSFWorkbook libro = new XSSFWorkbook();
    public XSSFSheet hoja;
    public String urlCreacion="";       //Es la direccion donde se guradara el libro INCLUIDO el nombre del archivo con su extension respectiva 
    ArrayList<Integer> filas = new  ArrayList<Integer>();       //lista de filas que contiene el libro
    ArrayList<Integer> columnas = new  ArrayList<Integer>();    //Lista de columnas que posee el libro
    
    WriteEcxel(String url){
        
    }
    
}

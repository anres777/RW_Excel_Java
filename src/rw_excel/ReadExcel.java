package rw_excel;

//Importaciones:
import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.BorderStyle;
 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;


public class ReadExcel {
    //es como un arreglo en 3 dimensiones, 
    // - la primer llave (tipo String) es el nombre de la hoja (contiene un mapa con el codigo de celda y su valor)
    // - la segunda llave es el codigo de la celda y el valor es el valor de la celda
    public Map<String, HashMap<String, String>> list_mapv = new HashMap<String, HashMap<String, String>>();
    
    ReadExcel(String rutaArchivo){
        DataFormatter formatter = new DataFormatter();        
        HashMap<String, String> map;
        try (FileInputStream file = new FileInputStream(new File(rutaArchivo))) {
            // leer archivo excel
            XSSFWorkbook worbook = new XSSFWorkbook(file);
            System.out.println("Numero de Hojas"+worbook.getNumberOfSheets());
            FormulaEvaluator evaluator = worbook.getCreationHelper().createFormulaEvaluator();
            //Inicio
            Iterator<Sheet> sheetIterator = worbook.iterator();
            
            while (sheetIterator.hasNext()) {
                //-- Obtener Hoja 
                Sheet sheet = sheetIterator.next();
                System.out.println("Hoja = "+ sheet.getSheetName());
                map = new HashMap<String, String>();
                Iterator<Row> rowIterator = sheet.iterator();
                Row row;    
                // se recorre cada fila hasta el final
                while (rowIterator.hasNext()) {
                    //--- Obtiene Fila
                    row = rowIterator.next(); 
                    Iterator<Cell> cellIterator = row.cellIterator();
                    Cell cell;
                    //se recorre cada celda
                    while (cellIterator.hasNext()) {
                        // se obtiene la celda en específico 
                        cell = cellIterator.next();
                        String contenidoCelda="";
                        switch (cell.getCellType()) 
                        {
                            case ERROR:  
                                System.out.println("ERROR EN LECTURA DE CELDA "+cell.getAddress().toString());
                                break;
                            case FORMULA:  
                                Cell nuevo = evaluator.evaluateInCell(cell);
                                switch (nuevo.getCellType()) {
                                    case ERROR:  
                                        System.out.println("ERROR EN LECTURA DE CELDA "+cell.getAddress().toString());
                                        break;
                                    case NUMERIC:  
                                        contenidoCelda= String.valueOf(nuevo.getNumericCellValue());
                                        break;
                                    case STRING:  
                                        contenidoCelda=nuevo.getStringCellValue();
                                        break;
                                }
                                break;
                            case NUMERIC:  
                                if( DateUtil.isCellDateFormatted(cell) ){
                                    Date fecha= cell.getDateCellValue();
                                    contenidoCelda = new SimpleDateFormat("dd/MM/yyyy").format(fecha);
                                }else{
                                    contenidoCelda= String.valueOf(cell.getNumericCellValue());
                                }
                                break;
                            case STRING:  
                                contenidoCelda=cell.getStringCellValue();
                                break;
                            case BOOLEAN:  
                                contenidoCelda = formatter.formatCellValue(cell);
                                break;
                            default: 
                                //dayString = "Dia inválido";
                                break;
                        }                        
                        if (contenidoCelda!= "") {
                            map.put(cell.getAddress().toString(), contenidoCelda);
                        }
                    }                                      
                }                
                this.list_mapv.put(sheet.getSheetName(), map); 
            }               
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }
    
    public String getCelda(String hoja, String codCelda){
        String res="";
        if (this.list_mapv.containsKey(hoja)) {
            if (this.list_mapv.get(hoja).containsKey(codCelda)) {
                res=this.list_mapv.get(hoja).get(codCelda);
            }else{
                System.out.println("la celda no existe"+codCelda);
            }
        }else{
            System.out.println("La hoja no existe"+hoja);
        }
        return res;
    }
    
}

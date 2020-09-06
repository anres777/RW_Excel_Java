
package rw_excel;

import java.text.ParseException;

/**
 *  Notas:
 *  - La urlSalida cambia de acuerdo al sistema operativo y tambien de acuerdo al nombre de usuario de la PC
 *  - Puedes escoger varios de los siguientes estilos para tus celdas.
 *      "decimal", "fecha", "negrita", "cursiva", "centrar", "alinear_derecha", "t_12", "t_14", "t_blanco", "t_rojo", "t_verde",
 *      "cafe", "rojo", "verde", "azul", "gris25%", "gris40%", "gris80%", "bordes"
 * @author Brayan
 */
public class RW_Excel {
    
    public static void main(String[] args) throws ParseException {
        //***** EJEMPLO DE CREACION DE LIBRO DE EXCEL ****                
//        String urlSalida="C:\\Users\\anres\\Desktop\\prueba.xlsx";
//        WriteEcxel hola = new WriteEcxel(urlSalida);
//        // Definiendo los estilos a utilizar en las celdas
//        String[] estiloTituloHoja = {"negrita","centrar","t_14","","",""};
//        String[] estiloTituloTabla = {"negrita","centrar","t_12","cafe","t_blanco","bordes"};
//        String[] estiloTexto = {"t_12","bordes"};
//        String[] estiloNumero = {"t_12","bordes","decimal"};
//        String[] estiloFecha = {"t_12","bordes","fecha"};
//        // Creando celdas
//        hola.crearCelda(0, 0, "Empresa X", estiloTituloHoja);
//        hola.combinarCentrar(0, 0, 2);        
//        hola.crearCelda(1, 0, "Fecha", estiloTituloTabla);
//        hola.crearCelda(1, 1, "Nombre", estiloTituloTabla);
//        hola.crearCelda(1, 2, "Sueldo", estiloTituloTabla);        
//        hola.crearCelda(2, 0, "4/10/2018", estiloFecha);
//        hola.crearCelda(2, 1, "Brayan Daniel Acosta", estiloTexto);
//        hola.crearCelda(2, 2, "12000", estiloNumero);        
//        hola.crearCelda(3, 0, "4/10/2020", estiloFecha);
//        hola.crearCelda(3, 1, "Eduardo Escobar", estiloTexto);
//        hola.crearCelda(3, 2, "15700", estiloNumero);        
//        hola.guardarLibro();

        //***** EJEMPLO DE LECTURA DE LIBRO DE EXCEL ******
//        ReadExcel hola = new ReadExcel("C:\\Users\\anres\\Desktop\\2018\\_CUENTAS POR COBRAR 2018.xlsx");
//        //El primer parametro es el nombre de la hoja, el segundo parametro es el codigo de la celda           
//        String prueba = hola.getCelda("Marco Antonio", "C8");
//        System.out.println(prueba);
          
    }
    
}

package mati.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
//import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.attribute.BasicFileAttributes;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFCell;


import java.io.FileOutputStream;


public class Main {

	
	/** 1)
	 * Cuando se tengan mas casos que se quieran procesar, solo generar una variable de la misma forma que aca abajo. 
	 * Que tenga la RUTA hacia la CARPETA QUE CONTIENE los archivos .DOCX de los alumnos. 
	 */
	
	static String rutaAlumnosMM = "MATRICULA2022/FICHA INCRIPCION MEDIO MAYOR 2022/MEDIO MAYOR 2022";
	static String rutaAlumnosK = "MATRICULA2022/FICHA INSCRIPCION KINDER 2022";
	static String rutaAlumnosPK = "MATRICULA2022/FICHA INSCRIPCION PREKINDER 2022/PREKINDER 2022";
	
	/** 2)
	 * Una vez se tenga la RUTA de la CARPETA CONTENEDORA. Crear una variable tal como esta hecho a continuacion. Y poner en el 
	 * parametro la RUTA anteriormente creada.
	 */
	static List<String> listaAlumnosMM = listarArchivosDocx(rutaAlumnosMM);
	static List<String> listaAlumnosK = listarArchivosDocx(rutaAlumnosK);
	static List<String> listaAlumnosPK = listarArchivosDocx(rutaAlumnosPK);

	
	static List<String> encabezados = new ArrayList<String>(Arrays.asList("NOMBRES", "APELLIDOS", "RUT", "DIRECCION",
			"TELEFONO", "COMUNA", "CODIGO", "AÑO MATRICULA", "DIA NAC", "MES NAC", "AÑO NAC", "OBSERVACIONES"));

	public static void main(String[] args) throws FileNotFoundException, IOException {

		
		/** 3)
		 *	Una vez se tengan esas dos variables anteriores hechas.  
		 *	pasarlas como parametro al metodo procesoFinal con su ruta y lista correspondiente. 
		 *	luego apretar click derecho en Main.java -> run as -> Java Application
		 */
		
		
		procesoFinal(rutaAlumnosMM, listaAlumnosMM);

		procesoFinal(rutaAlumnosK, listaAlumnosK);
		
		procesoFinal(rutaAlumnosPK, listaAlumnosPK);
		
	} 
	
	
	
	
	
	public static void procesoFinal( String rutaArchivos, List<String> listaAlumnos ) throws FileNotFoundException, IOException {
		

		for (int x = 0; x < 3; x++) {

			// Crea un libro de Excel en memoria
			SXSSFWorkbook workbook = new SXSSFWorkbook();

			// Crea una hoja dentro del libro
			SXSSFSheet sheet = workbook.createSheet("Hoja 1");

			// Crea la fila de t�tulos en la hoja
			SXSSFRow row = sheet.createRow(0);

			// Crea las celdas para cada t�tulo y agrega el t�tulo a cada celda
			for (int i = 0; i < encabezados.size(); i++) {
				SXSSFCell cell = row.createCell(i);
				cell.setCellValue(encabezados.get(i));
			}

				wordReader(rutaArchivos, listaAlumnos, sheet, workbook); // FUNCIONA

		}
		
	}
	

	public static void wordReader(String rutaArchivos, List<String> listaAlumnos, SXSSFSheet sheet,
			SXSSFWorkbook workbook) throws FileNotFoundException, IOException {

		int actualRow = 1;
		int contadorLista = 0;

		for (String archivoAlumno : listaAlumnos) {
			
			System.out.println("Leyendo y traspasando datos de archivo: "+archivoAlumno+ " a Excel\n");

			int celda = 0;
			int fila = 0;

			// SIRVE QUE SUBA CUANDO TOME OTRO NI�O

			List<SXSSFRow> listaFilas = new ArrayList<SXSSFRow>();


			listaFilas.add(sheet.createRow(actualRow));


			XWPFDocument doc = null;
			try (InputStream is = new FileInputStream(rutaArchivos + "/" + archivoAlumno)) {
				doc = new XWPFDocument(is);
			} catch (Exception e) {
				e.printStackTrace();
			}

			List<XWPFTable> tables = doc.getTables();

			for (XWPFTable table : tables) {

				for (XWPFTableRow row : table.getRows()) {

					//System.out.print("Fila " + fila + "= ");
					List<XWPFTableCell> cells = row.getTableCells();

					for (XWPFTableCell cell : cells) {

						//System.out.print("Celda " + celda + "= " + cell.getText() + "\t");

						if (fila == 1 && celda == 1) { // NOMBRES

							SXSSFCell celdaActual = listaFilas.get(contadorLista).createCell(0);
							celdaActual.setCellValue(cell.getText());

						} else if (fila == 2 && celda == 1) { // APELLIDOS

							SXSSFCell celdaActual = listaFilas.get(contadorLista).createCell(1);
							celdaActual.setCellValue(cell.getText());

						} else if (fila == 3 && celda == 1) { // RUT

							SXSSFCell celdaActual = listaFilas.get(contadorLista).createCell(2);
							celdaActual.setCellValue(cell.getText());

						} else if (fila == 4 && celda == 1) { // DIRECCION

							SXSSFCell celdaActual = listaFilas.get(contadorLista).createCell(3);
							celdaActual.setCellValue(cell.getText());

						} else if (fila == 6 && celda == 1) { // TELEFONO

							SXSSFCell celdaActual = listaFilas.get(contadorLista).createCell(4);
							celdaActual.setCellValue(cell.getText());

						} else if (fila == 7 && celda == 1) { // COMUNA

							SXSSFCell celdaActual = listaFilas.get(contadorLista).createCell(5);
							celdaActual.setCellValue(cell.getText());

						} else if (fila == 8 && celda == 1) { // CODIGO

							SXSSFCell celdaActual = listaFilas.get(contadorLista).createCell(6);
							celdaActual.setCellValue(getCodigoCurso(rutaArchivos));

						} else if (fila == 9 && celda == 1) { // A�O MATRICULA

							SXSSFCell celdaActual = listaFilas.get(contadorLista).createCell(7);
							celdaActual.setCellValue(getYearArchivoCarpeta(rutaArchivos));

						} else if (fila == 11 && celda == 1) { // FECHA NACIMIENTO CELDA 8, 9, 10 DIA MES A�O
							int j = 8;

							if (cell.getText().contains("/")) {

								String[] fechas = cell.getText().split("/");

								for (String fecha : fechas) {

									SXSSFCell celdaActual = listaFilas.get(contadorLista).createCell(j);
									celdaActual.setCellValue(fecha);
									j++;
								}

							} else if (cell.getText().contains("-")) {

								String[] fechas = cell.getText().split("-");

								for (String fecha : fechas) {

									SXSSFCell celdaActual = listaFilas.get(contadorLista).createCell(j);
									celdaActual.setCellValue(fecha);
									j++;
								}
							}

						} else if (fila == 16 && celda == 1) { // OBSERVACIONES

							SXSSFCell celdaActual = listaFilas.get(contadorLista).createCell(11);
							celdaActual.setCellValue(cell.getText());

						}

						celda++;

					}

					fila++;
					celda = 0;

					// System.out.println();

				}

			}

			actualRow++;

		} // FIN FOR EACH GRANDE

		contadorLista++;
		
		
		//Esto sirve para ajustar las columnas de excel al tama�o del encabezado menos la columna de Observaciones
		for (int i = 0; i < 10; i++) {		   
			sheet.trackColumnForAutoSizing(i);
			sheet.autoSizeColumn(i);
		}

		

		// Escribe el libro de Excel a un archivo en disco FINAL!!

		String nombreArchivo = "";

		if (getCodigoCurso(rutaArchivos) == "MM")
			nombreArchivo = "Medio Mayor Nomina " + getYearArchivoCarpeta(rutaArchivos);
		else if (getCodigoCurso(rutaArchivos) == "PK")
			nombreArchivo = "PreKinder Nomina " + getYearArchivoCarpeta(rutaArchivos);
		else if (getCodigoCurso(rutaArchivos) == "K")
			nombreArchivo = "Kinder Nomina " + getYearArchivoCarpeta(rutaArchivos);

		try (FileOutputStream fos = new FileOutputStream("INFORME2022/" + nombreArchivo + ".xlsx")) {
			workbook.write(fos);
			System.out.println("--------------------------- \n LISTO EXCEL: "+ nombreArchivo + "\n---------------------------");
		}

	} // Fin wordReader

	
	
	public static String getCodigoCurso(String rutaArchivo) {

		String codigo = "";

		if (rutaArchivo.contains("MEDIO MAYOR")) 
			codigo = "MM";
		else if (rutaArchivo.contains("PRE") && rutaArchivo.contains("KINDER"))
			codigo = "PK";
		else if (!rutaArchivo.contains("PRE") && rutaArchivo.contains("KINDER")) 
			codigo = "K";

		return codigo;
	}

	public static String getYearArchivoCarpeta(String rutaArchivo) {

		String anio = rutaArchivo.substring(rutaArchivo.length() - 4, rutaArchivo.length());

		// System.out.println("El a�o del archivo es: "+ anio);
		return anio;
	}

	public static List<String> listarArchivosDocx(String ruta) { // List<File[]>

		File[] listaAlumnos = new File(ruta).listFiles((dir, name) -> name.endsWith(".docx"));

		List<String> listaAlumnosString = new ArrayList<String>();

		// Recorre cada archivo .docx y imprime su nombre
		// System.out.println(" \n Lista " + getCodigoCurso(ruta));
		for (File docxFile : listaAlumnos) {
			listaAlumnosString.add(docxFile.getName());
		}

		// System.out.println(listaAlumnosString);
		return listaAlumnosString;

	}

	public static int getYearArchivoMetadato(String rutaArchivo, String nombreArchivo) throws IOException {

		// Crea una instancia de Path para el archivo
		Path file = Paths.get(rutaArchivo + "/" + nombreArchivo);

		BasicFileAttributes attrs = Files.readAttributes(file, BasicFileAttributes.class);

		Instant creationTime = attrs.creationTime().toInstant();

		ZonedDateTime zonedDateTime = creationTime.atZone(ZoneId.systemDefault());

		LocalDateTime dateTime = zonedDateTime.toLocalDateTime();

		int year = dateTime.getYear();

		System.out.println("El a�o de creaci�n del archivo es: " + year);

		return year;

	}

}

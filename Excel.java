package davi.daviplata.nacional.android.utilidades;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Date;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.Row;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cucumber.api.Scenario;
import net.sourceforge.htmlunit.corejs.javascript.Evaluator;

import org.apache.poi.ss.usermodel.Sheet;

public class Excel {

	XSSFWorkbook workbookEvidencias = new XSSFWorkbook();
	XSSFSheet Datos;
	XSSFSheet hojaEvidencia;
	XSSFRow filaEvidencia;
	XSSFRow celdaColumnA;
	String[] header;
	String[] datos;
	String[] scenarioo;
	
	String rutaArchivoTempo = "";
	File rutaArchivoEvidencia;
	Cell c;
	static int contadorFila = 1;
	int contadorCelda = 1;
	BaseUtil base;

	// Anadir elementos
	XSSFSheet hoja1 = null;
	XSSFWorkbook wb = null;
	BufferedImage in;
	Scenario scenario;

	public Excel(BaseUtil base) {
		super();
		this.base = base;
		scenarioo = base.scenario.getName().split("_");

	}

	// Metodo para crear o leer un archivo
	public void crearLeerArchivo(String rutaEvidencia, String nombreArchivo) {
		rutaArchivoTempo = rutaEvidencia + "/" + nombreArchivo;
		rutaArchivoEvidencia = new File(rutaArchivoTempo);
		if (rutaArchivoEvidencia.exists()) {
			datos = datos();
			buscarFilas();
			System.out.println("Valor de contador inicial"+ contadorFila);
			ingresarDatos(contadorFila, hoja1);
			autoSize(hoja1, contadorFila);
			try {
				addImage();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			try {
				escribirExcel(rutaArchivoTempo, wb);
			} catch (IOException e) {
				System.out.println("No se cerro el archivo");
			}
		} else {

			System.out.println("Creo el archivo");
			workbookEvidencias = new XSSFWorkbook();
			Datos = workbookEvidencias.createSheet("Datos");
			hojaEvidencia = workbookEvidencias.createSheet("Evidencia");
			header = header();
			cabecera();
			autoSize(Datos, 2);
			datos = datos();

			try {
				cerrarExcel(rutaArchivoTempo, workbookEvidencias);// Creo un libro de excel en blanco
			} catch (Exception e) {
				System.out.println(e);
			}
			
			buscarFilas();
			ingresarDatos(contadorFila, hoja1);
			autoSize(hoja1, contadorFila);
			try {
				addImage();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			try {
				escribirExcel(rutaArchivoTempo, wb);
			} catch (IOException e) {
				System.out.println("No se cerro el archivo");
			}
		}
	}

	// Escribe el excel
	public void escribirExcel(String ruta, XSSFWorkbook wb) throws IOException {
		FileOutputStream fileOut = null;
		fileOut = new FileOutputStream(ruta);
		wb.write(fileOut);
		fileOut.close();
	}

	// Cierra el excel
	public void cerrarExcel(String ruta, XSSFWorkbook wb) throws IOException {
		FileOutputStream auxiliarEscritura = new FileOutputStream(ruta);
		wb.write(auxiliarEscritura);
		auxiliarEscritura.flush();
		auxiliarEscritura.close();
		workbookEvidencias.close();
	}

	// Datos de encabezados
	public String[] header() {
		String[] header = new String[] { "", "Numero de caso","ID Usuario", "Saldo inicial Daviplata", "Num Cel o Cuenta",
				"Monto de transaccion", "Saldo Final Daviplata", "ID de transaccion" };
		return header;
	}

	// Datos de prueba
	public String[] datos() {
		String[] datos = new String[] { "",scenarioo[0], base.usuario, base.saldoIni, base.cuentaONumero+"-"+base.cuentaONumero2, base.monto, base.saldoFin,
				base.idTransaccion };
		return datos;
	}

	// Metodo para insertar la cabecera en el excel
	public void cabecera() {
		for (int i = 1; i < header.length; i++) {
			XSSFRow row = Datos.createRow(i);
			for (int j = 1; j < header.length; j++) {
				if (i == 1) {
					XSSFCell cell = row.createCell(j);
					cell.setCellStyle(styleCabecera(workbookEvidencias));
					cell.setCellValue(header[j]);
				}
			}
		}
	}

	// Estilo de la cabecera
	public CellStyle styleCabecera(XSSFWorkbook wb) {
		CellStyle style = wb.createCellStyle();
		Font font = wb.createFont();
		font.setBold(true);
		style.setFont(font);
		return style;

	}

	// Ajustar las celdas
	public void autoSize(XSSFSheet sh, int fila) {
		for (int i = 0; i <= sh.getRow(fila).getPhysicalNumberOfCells(); i++) {
			sh.autoSizeColumn(i);
		}
	}

	// Escribir en el excel
	public void ingresarDatos(int fila, XSSFSheet sf) {
		XSSFRow row = sf.createRow(fila);
		for (int j = 1; j < datos.length; j++) {
			XSSFCell cell = row.createCell(j);
			cell.setCellValue(datos[j]);

		}
	}

	// Busco la fila disponible para escribir la data
	public void buscarFilas() {
		try (FileInputStream file = new FileInputStream(rutaArchivoEvidencia)) {
			wb = new XSSFWorkbook(file);
			hoja1 = wb.getSheetAt(0);
			for (Sheet sheet : wb) {
				for (Row row : sheet) {
					isRowEmpty(row);
				}

			}

		} catch (Exception e) {

		}

	}

	// Metodo para determinar que celdas estan vacias
	public static boolean isRowEmpty(Row row) {
		for (int c = row.getFirstCellNum(); c < 1000; c++) {
			Cell cell = row.getCell(c);
			if (cell != null && cell.getCellType() != CellType.BLANK) {
				String value = cell.getStringCellValue();
				System.out.println(value);
				if(value.contains("SYS")) {
					return false;
				}else {
					contadorFila++;
					return false;
				}
				
				

				
			}

		}
		return true;

	}

	public void addImage() throws IOException {
		int t = 5;
		int col = 4;
		int row = 2;
		int i = 0;
		int f = 0;
		int espacio = 25;
		int casoFila = 1;
		double scaleX = 0.0;
		double scaleY = 0.0;
		
		hojaEvidencia = wb.getSheetAt(1);
		// Agregar el caso al excel
		// Agregar mas imagenes
		// XSSFRow Roww = hojaEvidencia.createRow(1);
		// Cell celll = Roww.createCell(1);
		System.out.println("El contador es " + contadorFila);
		if (contadorFila == 2) {
			System.out.println("Entro 1");
			XSSFRow Row = hojaEvidencia.createRow(1);
			Cell cell = Row.createCell(4);
			cell.setCellValue(base.NombreSce);
			row = 4;
		} else if (contadorFila == 3) {
			XSSFRow Row = hojaEvidencia.createRow(espacio - 2);
			Cell cell = Row.createCell(4);
			cell.setCellValue(base.NombreSce);
			System.out.println("Entro 2");
			row = espacio;
		} else {
			System.out.println("Entre 3");
			contadorFila = contadorFila - 2;
			System.out.println("El contador es en  " + contadorFila);
			row = espacio * contadorFila;
			XSSFRow Row = hojaEvidencia.createRow(row - 2);
			Cell cell = Row.createCell(4);
			cell.setCellValue(base.NombreSce);
		}
		do {
			//Cambiar PNG a JPG/JPEG cuando se requiera ****
			InputStream inputStream = new FileInputStream(System.getProperty("user.dir") + "//Evidencias//"
					+ scenarioo[0] + "//" + base.NombreImage[i] + ".PNG");
			System.out.println("Nombre imagen " + base.NombreImage[i]);
			byte[] bytes = IOUtils.toByteArray(inputStream);
			int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
			inputStream.close();
			CreationHelper helper = wb.getCreationHelper();
			Drawing<?> drawing = hojaEvidencia.createDrawingPatriarch();
			ClientAnchor anchor = helper.createClientAnchor();
			anchor.setCol1(col);
			System.out.println("Valor de la fila " + row);
			anchor.setRow1(row);
			Picture pict = drawing.createPicture(anchor, pictureIdx);
				scaleX = t * 1.0 * 0.5908494;
				scaleY = t * 1.8 * 1.9677996;
				if(base.NombreImage[i].contains("web")) {
					pict.resize(10,15);
					col = col + 11;
				}else{
					pict.resize(scaleX,scaleY);
					col = col + 3;
				}
				
				
			i++;
		} while (base.NombreImage[i] != null);
		contadorFila = 1 ;
	}

}

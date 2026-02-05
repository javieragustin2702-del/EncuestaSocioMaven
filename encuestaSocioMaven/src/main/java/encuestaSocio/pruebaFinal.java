package encuestaSocio;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import encuestaSocio.funciones;
import org.apache.poi.ss.usermodel.*;

public class pruebaFinal {

	public static void main(String[] args) {
		/*
		 * El localDate de fecha(línea 20) hay que cambiarlo a 6 si se realiza en el
		 * mismo año de los datos que se quieren meter en el excel a parte de cambiar
		 * ese localDate hay que cambiar los minusYears de las líneas 30/31 y 104/105 a
		 * 5 y 0 respectivamente para que cree el archivo con el nombre de los años que
		 * entran en el excel.
		 * 
		 * La idea es que busca el año que se quiere eliminar. Para ello mueve lo que
		 * tiene al lado y la columna que queda vacía se rellena con los datos de este
		 * año. Para obtener los datos toma como referencia en EncuestaSocio la celda B2
		 * y la celda anterior en la que se trabajó y, si es en la misma fila que la vez
		 * anterior,cogerá los datos de la columna de la derecha y si es en otra fila
		 * coge los datos de 8 filas más abajo. Esto solo cambia con las filas 53 y y
		 * 114 que en el caso de la línea 53 en vez de coger los datos de la columna de
		 * la derecha,coge los de 3 columnas a la derecha ya que los datos de
		 * "no lo puso" están en EncuestaSocio 3 columanas a la derecha con respecto a
		 * la celda de "primera posición" y la línea 114 de EncuestaSocio los datos que
		 * se necesitan están 8 filas más abajo todo el rato por lo que a partir de esa
		 * línea el programa rellena con los datos de 8 filas más abajo,
		 * indiferentemente de que en EvolucionEncuestaSocio se este trabajando en la
		 * misma línea que la vez anterior. En caso de que se modifique la estructura de
		 * EncuestaSocio, el excel resultante de la evolución de los datos con los años
		 * saldrá mal.
		 */
		LocalDate fecha = LocalDate.now().minusYears(7);
		LocalDate hoy = LocalDate.now();
		int año = fecha.getYear();

		// Archivos con los que trabaja
		File archivoPlantilla = new File("src/main/java/excels/EvolucionEncuestaSocio.xlsx");
		File archivoOrigen = new File("src/main/java/excels/encuestaSocio.xlsx");

		// Archivo resultante
		File archivoResultado = new File("src/main/java/excels/EvolucionEncuestaSocio " + hoy.minusYears(6).getYear()
				+ "-" + hoy.minusYears(1).getYear() + ".xlsx");

		try (Workbook wbDestino = WorkbookFactory.create(new FileInputStream(archivoPlantilla));
				Workbook wbOrigen = WorkbookFactory.create(new FileInputStream(archivoOrigen))) {
			// trabaja con la 1ª hoja de cada excel
			Sheet sheetDestino = wbDestino.getSheetAt(0);
			Sheet sheetOrigen = wbOrigen.getSheetAt(0);

			// Punto de origen de EncuestaSocio
			int filaOrigenBase = 1; // B2
			int colOrigenBase = 1; // columna B

			Integer filaDestinoAnterior = null;

			// ===== CONTROL ESPECÍFICO FILA 53 =====
			int contadorFila53 = 0;
			int colBaseFila53 = 1; // Columna 1 hace referencia a la columna B

			for (int i = 0; i <= sheetDestino.getLastRowNum(); i++) {

				Row row = sheetDestino.getRow(i);
				if (row == null)
					continue;

				for (Cell cell : row) {

					boolean esAño = false;

					if (cell.getCellType() == CellType.NUMERIC && (int) cell.getNumericCellValue() == año) {
						esAño = true;
					}

					if (!esAño)
						continue;

					int filaActual = cell.getRowIndex();

					if (filaDestinoAnterior != null) {
						// Condición para cuando llega a la fila 53 en EvoluciónEncuestaSocio
						if (filaActual == 52) {
							filaOrigenBase = 41;
							colOrigenBase = colBaseFila53 + (contadorFila53 * 3);
							contadorFila53++;
							// Condición para cuando llega a la fila 114 en EncuestaSocio
						} else if (filaActual >= 114) {
							filaOrigenBase += 8;
							colOrigenBase = 1;
							// Condición si la fila en la que esta trabajando es la misma que la última vez
						} else if (filaActual == filaDestinoAnterior) {
							colOrigenBase++;
							// Condición si la fila en la que esta trabajando no es la misma que la última
							// vez
						} else {
							filaOrigenBase += 8;
							colOrigenBase = 1;
						}
					}

					// Funciones que realiza al encontrar el año que se quiere borrar. Están en
					// funciones
					funciones.moverBloqueIzquierda(sheetDestino, cell);
					funciones.rellenar2025YDatos(sheetDestino, cell, sheetOrigen, filaOrigenBase, colOrigenBase);
					filaDestinoAnterior = filaActual;
				}
			}

			// Rellena el excel resultante
			try (FileOutputStream fos = new FileOutputStream(archivoResultado)) {
				wbDestino.write(fos);
			}

			System.out.println("EvolucionEncuestaSocio " + hoy.minusYears(6).getYear() + "-"
					+ hoy.minusYears(1).getYear() + ".xlsx hecho");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}

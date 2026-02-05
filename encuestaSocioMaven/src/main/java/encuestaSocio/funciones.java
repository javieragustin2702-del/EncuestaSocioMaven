package encuestaSocio;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class funciones {

	// ================= MOVER BLOQUE =================

	public static void moverBloqueIzquierda(Sheet hoja, Cell cell) {

		int colBase = cell.getColumnIndex();
		int filaBase = cell.getRowIndex();

		for (int fila = filaBase; fila <= filaBase + 6; fila++) {

			Row row = hoja.getRow(fila);
			if (row == null)
				continue;

			for (int col = colBase + 1; col <= colBase + 5; col++) {

				Cell origen = row.getCell(col);
				if (origen == null)
					continue;

				Cell destino = row.getCell(col - 1);
				if (destino == null)
					destino = row.createCell(col - 1);

				copiarValorCelda(origen, destino);
				origen.setBlank();
			}
		}
	}

	// ================= COPIAR CELDA =================

	public static void copiarValorCelda(Cell origen, Cell destino) {

		switch (origen.getCellType()) {

		case STRING:
			destino.setCellValue(origen.getStringCellValue());
			break;

		case NUMERIC:
			if (DateUtil.isCellDateFormatted(origen)) {
				destino.setCellValue(origen.getDateCellValue());
			} else {
				destino.setCellValue(origen.getNumericCellValue());
			}
			break;

		case BOOLEAN:
			destino.setCellValue(origen.getBooleanCellValue());
			break;

		case FORMULA:
			destino.setCellFormula(origen.getCellFormula());
			break;

		default:
			destino.setBlank();
		}
	}

	// ================= RELLENAR 2025 + DATOS =================

	public static void rellenar2025YDatos(Sheet destino, Cell celdaBase, Sheet origen, int filaOrigen, int colOrigen) {

		int filaBase = celdaBase.getRowIndex();
		int col2025 = celdaBase.getColumnIndex() + 5;

		// ---- Escribir 2025 ----
		Row fila = destino.getRow(filaBase);
		if (fila == null)
			fila = destino.createRow(filaBase);

		Cell celda2025 = fila.getCell(col2025);
		if (celda2025 == null)
			celda2025 = fila.createCell(col2025);

		celda2025.setCellValue(2025);

		// ---- Copiar las 6 celdas debajo ----
		for (int i = 1; i <= 6; i++) {

			Row filaDestino = destino.getRow(filaBase + i);
			if (filaDestino == null)
				filaDestino = destino.createRow(filaBase + i);

			Cell celdaDestino = filaDestino.getCell(col2025);
			if (celdaDestino == null)
				celdaDestino = filaDestino.createCell(col2025);

			Row filaOrigenExcel = origen.getRow(filaOrigen + i - 1);
			if (filaOrigenExcel == null)
				continue;

			Cell celdaOrigen = filaOrigenExcel.getCell(colOrigen);
			if (celdaOrigen != null) {
				copiarValorCelda(celdaOrigen, celdaDestino);
			}
		}
	}
}

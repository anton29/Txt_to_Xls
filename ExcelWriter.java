
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;




public class ExcelWriter {

	private static final String folder = "C:\\Users\\U1\\Desktop\\Trip";

	private static final String importFileName = "JsonExample2.JSON";
	private static final String importPath = folder + "\\"+ importFileName;

	private static final String exportFileName = "JsonExample.xlsx";
	private static final String exportPath = folder + "\\"+ exportFileName;

	private static final String sheetName = "test";

	private static final String[] columns = {"Date/Time", "latitude", "longitude"};

	public static void main(String[] args) throws IOException, InvalidFormatException, ParseException {
		// Create a Workbook
		Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

		// Create a Sheet
		Sheet sheet = workbook.createSheet(sheetName);

		JSONParser parser = new JSONParser();

		// Create a Font for styling header cells
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.RED.getIndex());

		// Create a CellStyle with the font
		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);

		// Create a Row
		Row headerRow = sheet.createRow(0);

		// Create cells
		for(int i = 0; i < columns.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(columns[i]);
			cell.setCellStyle(headerCellStyle);
		}


		//open Json file
		Object obj = parser.parse(new FileReader(
				importPath));

		JSONArray entries = (JSONArray) obj;

		int rowNum = 1;
		for (int i = 0; i < entries.size(); i++) {
			Row row = sheet.createRow(rowNum++);

			JSONObject entry = (JSONObject) entries.get(i);
			String date_time = (String) entry.get("Date/Time ");
			JSONArray lat_long = (JSONArray) entry.get("Lat/Long ");

			row.createCell(0).setCellValue(date_time);
			row.createCell(1).setCellValue((double) lat_long.get(0));
			row.createCell(2).setCellValue((double) lat_long.get(1));
		}

		// Resize all columns to fit the content size
		for(int i = 0; i < columns.length; i++) {
			sheet.autoSizeColumn(i);
		}

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream(exportPath);
		workbook.write(fileOut);
		fileOut.close();

		// Closing the workbook
		workbook.close();
	}
}

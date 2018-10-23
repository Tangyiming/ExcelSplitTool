package application;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import javafx.stage.Stage;

public class Controller {

	@FXML
	Button OpenFile;
	@FXML
	TextField FilePath;
	@FXML
	ComboBox<String> ChooseSheet;
	@FXML
	TextField TitleCount;
	@FXML
	TextField SplitCount;
	@FXML
	Button Split;
	private static Stage newAlertDialog;
	InputStream input = null;
	Workbook wb = null;
	List<String> l = new ArrayList<String>();
	String abpath;

	@FXML
	public void openfile(ActionEvent event) throws Exception {
		FileChooser fileChooser = new FileChooser();
		fileChooser.setTitle("选择拆分excel文件");
		fileChooser.getExtensionFilters().addAll(new ExtensionFilter("Excel 工作薄", "*.xls", "*.xlsx"),
				new ExtensionFilter("所有文件", "*.*"));
		File selectedFile = fileChooser.showOpenDialog(newAlertDialog);

		if (selectedFile != null) {
			FilePath.setText(selectedFile.getAbsolutePath());
		}

		abpath = FilePath.getText();

		try {
			input = new FileInputStream(abpath);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}

		if (abpath.endsWith(".xlsx")) {
			wb = new XSSFWorkbook(input);
		} else {
			wb = new HSSFWorkbook(input);
		}
		int sn = wb.getNumberOfSheets();
		for (int i = 0; i < sn; i++) {
			String sheetname = wb.getSheetName(i);
			l.add(sheetname);
		}
		ChooseSheet.getItems().addAll(l);
	}

	@SuppressWarnings("deprecation")
	@FXML
	public void split() throws IOException {
		Sheet st;
		String cc;
		int tc = Integer.parseInt(TitleCount.getText());
		int sc = Integer.parseInt(SplitCount.getText());
		int index = tc-1;
		if (ChooseSheet.getValue() != null) {
			st = wb.getSheet(ChooseSheet.getValue());
		} else {
			st = wb.getSheetAt(0);
		}
		int ar = st.getPhysicalNumberOfRows();
		int times = 0;
		while (index < ar) {
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet();
			for (int i = 0; i < tc; i++) {
				Row r = st.getRow(i);
				if (r != null) {
					HSSFRow row = sheet.createRow(i);
					for (Cell cell : r) {
						int cn = cell.getColumnIndex();

						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_STRING:
							cc = cell.getRichStringCellValue().getString();
							break;
						case Cell.CELL_TYPE_NUMERIC:
							if (DateUtil.isCellDateFormatted(cell)) {
								cc = String.valueOf(cell.getDateCellValue());
							} else {
								cc = String.valueOf(cell.getNumericCellValue());
							}
							break;
						case Cell.CELL_TYPE_BOOLEAN:
							cc = String.valueOf(cell.getBooleanCellValue());
							break;
						case Cell.CELL_TYPE_FORMULA:
							cc = String.valueOf(cell.getCellFormula());
							break;
						case Cell.CELL_TYPE_ERROR:
							cc = String.valueOf(cell.getErrorCellValue());
							break;
						default:
							cc = "";
						}
						cell = row.createCell(cn);
						cell.setCellType(HSSFCell.CELL_TYPE_STRING);
						cell.setCellValue(cc);
					}
				}
			}

			for (int j =index+1 ; j < index + sc+1; j++) {
				Row r = st.getRow(j);
				if (r != null) {
					HSSFRow row = sheet.createRow(j-index+1);
					for (Cell cell : r) {
						int cn = cell.getColumnIndex();
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_STRING:
							cc = cell.getRichStringCellValue().getString();
							break;
						case Cell.CELL_TYPE_NUMERIC:
							if (DateUtil.isCellDateFormatted(cell)) {
								cc = String.valueOf(cell.getDateCellValue());
							} else {
								cc = String.valueOf(cell.getNumericCellValue());
							}
							break;
						case Cell.CELL_TYPE_BOOLEAN:
							cc = String.valueOf(cell.getBooleanCellValue());
							break;
						case Cell.CELL_TYPE_FORMULA:
							cc = String.valueOf(cell.getCellFormula());
							break;
						case Cell.CELL_TYPE_ERROR:
							cc = String.valueOf(cell.getErrorCellValue());
							break;
						default:
							cc = "";
						}
						cell = row.createCell(cn);
						cell.setCellType(HSSFCell.CELL_TYPE_STRING);
						cell.setCellValue(cc);
					}
				}
			}
			times+=1;
			index+=sc;
			System.out.println("after:"+index);
			File file = new File(abpath.trim());
			String fileDirectory = file.getParent();
			String originfilename = file.getName().split("\\.")[0];
			FileOutputStream fo = new FileOutputStream(fileDirectory + "\\" + originfilename + "_" + times + ".xls");
			workbook.write(fo);
			fo.close();
			workbook.close();
		}
	}
}

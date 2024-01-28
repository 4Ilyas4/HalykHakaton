package calculator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Calculator {
	final static String filePath = "src/main/java/calculator/calc.xlsx";
	static String loss_statistics = "нет";//
	public double cost;
	public double prem;

	public static void writeDataWithPOI(String insurer, String bin, String insurerbefore, String oked,
										String sign_of_affiliation, String region, String status_of_contract, String reinsurer,
										String date_of_conclusion, int workers_number, double declared_gfzp, double declared_sum_insured,
										String loss_statistics, String periodicity_of_insurance_premium) {
		try (FileInputStream fileInputStream = new FileInputStream(filePath);

			 Workbook workbook = new XSSFWorkbook(fileInputStream)) {

			Sheet sheet = workbook.getSheet("tpt");
			if (sheet == null) {
				sheet = workbook.createSheet("tpt");
			}

			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			sheet.setForceFormulaRecalculation(true);

			// Получение или создание строки
			Row row = sheet.getRow(7);
			if (row == null) {
				row = sheet.createRow(7);
			}

			// Запись данных
			row.createCell(0).setCellValue(insurer);
			row.createCell(1).setCellValue(bin);
			row.createCell(2).setCellValue(insurerbefore);

			row = sheet.getRow(11);
			if (row == null) {
				row = sheet.createRow(11);
			}
			row.createCell(0).setCellValue(oked);

			row = sheet.getRow(14);
			if (row == null) {
				row = sheet.createRow(14);
			}
			row.createCell(0).setCellValue(sign_of_affiliation);
			row.createCell(1).setCellValue(region);

			row = sheet.getRow(17);
			if (row == null) {
				row = sheet.createRow(17);
			}
			row.createCell(0).setCellValue(status_of_contract);
			row.createCell(1).setCellValue(reinsurer);

			row = sheet.getRow(19);
			if (row == null) {
				row = sheet.createRow(19);
			}
			row.createCell(1).setCellValue(date_of_conclusion);

			row = sheet.getRow(21);
			if (row == null) {
				row = sheet.createRow(21);
			}
			row.createCell(1).setCellValue(workers_number);

			row = sheet.getRow(24);
			if (row == null) {
				row = sheet.createRow(24);
			}
			row.createCell(0).setCellValue(declared_gfzp);
			row.createCell(2).setCellValue(declared_sum_insured);

			row = sheet.getRow(28);
			if (row == null) {
				row = sheet.createRow(28);
			}
			row.createCell(1).setCellValue(loss_statistics);

			row = sheet.getRow(36);
			if (row == null) {
				row = sheet.createRow(36);
			}
			row.createCell(2).setCellValue(periodicity_of_insurance_premium);

			// Для перерасчета формул
			evaluator.evaluateAll();
			try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
				workbook.write(fileOut);
			}

		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public double readNumericCellValue(Cell cell, Workbook workbook) {
		if (cell.getCellType() == CellType.NUMERIC) {
			return cell.getNumericCellValue();
		} else if (cell.getCellType() == CellType.FORMULA) {
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			CellValue cellValue = evaluator.evaluate(cell);
			return cellValue.getNumberValue();
		}
		return 0.0;
	}

	public void readDataWithPOI() {
		try (FileInputStream fileInputStream = new FileInputStream(filePath);
			 Workbook workbook = new XSSFWorkbook(fileInputStream)) {

			Sheet sheet = workbook.getSheet("tpt");

			// Чтение Цены
			Cell costCell = sheet.getRow(24).getCell(4);
			if (costCell.getCellType() == CellType.FORMULA) {
				FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
				CellValue cellValue = evaluator.evaluate(costCell);
				cost = cellValue.getNumberValue();
			} else if (costCell.getCellType() == CellType.NUMERIC) {
				cost = costCell.getNumericCellValue();
			}

			// Чтение Премии
			Cell premCell = sheet.getRow(36).getCell(1);
			if (premCell.getCellType() == CellType.FORMULA) {
				FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
				CellValue cellValue = evaluator.evaluate(premCell);
				prem = cellValue.getNumberValue();
			} else if (premCell.getCellType() == CellType.NUMERIC) {
				prem = premCell.getNumericCellValue();
			}

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public double[] getCostPrem(String insurer, String bin, String insurerbefore, String oked,
								String sign_of_affiliation, String region, String status_of_contract, String reinsurer,
								String date_of_conclusion, int workers_number, double declared_gfzp, double declared_sum_insured,
								String periodicity_of_insurance_premium) {
		writeDataWithPOI(insurer, bin, insurerbefore, oked, sign_of_affiliation, region, status_of_contract, reinsurer, date_of_conclusion, workers_number, declared_gfzp, declared_sum_insured, loss_statistics, periodicity_of_insurance_premium);
		try {
			Thread.sleep(300);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		readDataWithPOI();
		return new double[]{cost, prem};
	}
//	public static void main(String[] args) {
//		Calculator calculator = new Calculator();
//
//		// пример
//		String insurer = "";
//		String bin = "951140000042";
//		String insurerbefore = "АО КСЖ \\\"KM Life\\\"";
//		String oked = "47111";
//		String sign_of_affiliation = "да";
//		String region = "г. Алматы";
//		String status_of_contract = "входящее перестрахование";
//		String reinsurer = "Нет перестрахователя";
//		String date_of_conclusion = "19.09.2023";
//		int workers_number = 10;
//		double declared_gfzp = 5000.00;
//		double declared_sum_insured = 63300.00;
//		String periodicity_of_insurance_premium = "Единовременно";
//
//		// Write data to the Excel file
//		writeDataWithPOI(insurer, bin, insurerbefore, oked, sign_of_affiliation, region, status_of_contract, reinsurer, date_of_conclusion, workers_number, declared_gfzp, declared_sum_insured, loss_statistics, periodicity_of_insurance_premium);
//
//		try {
//			Thread.sleep(300);
//		} catch (InterruptedException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//		// Read data from the Excel file
//		calculator.readDataWithPOI();
//		// Output the results (optional, for verification)
//		System.out.println("Cost from Excel: " + calculator.cost);
//		System.out.println("Prem from Excel: " + calculator.prem);
//	}
}


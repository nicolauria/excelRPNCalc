// Excel RPN Calculator
// Reads xls/xlsx file (line 31)
// Writes to new xls/xlsx file (line 125)

package excelRPNCalc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Stack;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

public class excelRPNCalc {
	
	private static Workbook wb;
	private static Sheet sh;
	private static FileInputStream fis;
	private static FileOutputStream fos;
	private static Row row;
	private static Cell cell;
	
	public static void main(String[] args) throws Exception {
		String inputFile = args[0];
		fis = new FileInputStream(inputFile);
		wb = WorkbookFactory.create(fis);
		sh = wb.getSheet("Sheet1");
		
		DataFormatter formatter = new DataFormatter();
		String alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		boolean unresolvedCells = true;
		boolean unresolvedCell = false;
		
		while (unresolvedCells) {
			unresolvedCells = false;
			Iterator<Row> rowIterator = sh.rowIterator();
			
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				
				while (cellIterator.hasNext()) {
					Cell currentCell = cellIterator.next();
	                String cellValue = formatter.formatCellValue(currentCell);
	                if (isNumeric(cellValue)) continue;
	            
	                Stack<Double> stack = new Stack<>();
	                
	                unresolvedCell = false;
	                outerLoop:
	                for(String token : cellValue.split(" ")) {
	                	double secondOperand = 0.0;
	                    double firstOperand = 0.0;
	                	
	                	switch (token) {
	                		case "+":
	                			secondOperand = stack.pop();
	                            firstOperand = stack.pop();

	                            stack.push(firstOperand + secondOperand);
	                            break;
	                		case "-":
	                			secondOperand = stack.pop();
	                            firstOperand = stack.pop();

	                            stack.push(firstOperand - secondOperand);
	                			break;
	                		case "*":
	                			secondOperand = stack.pop();
	                            firstOperand = stack.pop();

	                            stack.push(firstOperand * secondOperand);
	                			break;
	                		case "/":
	                			secondOperand = stack.pop();
	                            firstOperand = stack.pop();

	                            stack.push(firstOperand / secondOperand);
	                			break;
	                		case "^":
	                			secondOperand = stack.pop();
	                            firstOperand = stack.pop();

	                            stack.push(Math.pow(firstOperand, secondOperand));
	                			break;
	                		default:
	                			char el = token.charAt(0);
	    	                	if (Character.isLetter(el)) {
	    	        				int rowNum = alphabet.indexOf(el);
	    	        				int colNum = Character.getNumericValue(token.charAt(1));
	    	        				
	    	        				Row row2 = sh.getRow(rowNum);
	    	        				Cell otherCell = row2.getCell(colNum - 1);
	    	        				String otherCellVal = formatter.formatCellValue(otherCell);
	    	           
	    	        				if (isNumeric(otherCellVal)) {
	    	        					stack.push(Double.parseDouble(otherCellVal));
	    	        				} else {
	    	        					unresolvedCells = true;
	    	        					unresolvedCell = true;
	    	        					break outerLoop;
	    	        				}
	    	        			} else {
	    	        				stack.push(Double.parseDouble(token));
	    	        			}
	                            break;
	                	}
	                	
	                	
	                }
	                if (stack.size() > 0 && unresolvedCell == false) {
	                	currentCell.setCellValue(stack.pop());
	                }
	            }
			}
		}
		fis.close();
		File outputFile = new File("./result.xlsx");
		FileOutputStream output_file = new FileOutputStream(outputFile);
		wb.write(output_file);
		output_file.close();
		System.out.println("Results written to file: " + outputFile);
	}
	
	public static boolean isNumeric(String strNum) {
	    try {
	        double d = Double.parseDouble(strNum);
	    } catch (NumberFormatException | NullPointerException nfe) {
	        return false;
	    }
	    return true;
	}
}
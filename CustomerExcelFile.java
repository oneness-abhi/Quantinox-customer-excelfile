package com.zidoka.customers.controller;

import java.io.FileOutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.bind.annotation.PutMapping;

import com.zidoka.customers.api.CustomerService;
import com.zidoka.customers.core.Customer;
import com.zidoka.customers.manager.CustomerManager;
@Service
public class CustomerExcelFile {
	
	@Autowired
	private CustomerManager customerManager;
	
	@PutMapping("/syncup")
    public String syncup(){
        long totalEnquiries = customerManager.syncup();
        return totalEnquiries +" records successfully synced up.";
    }
	public static void main(String[] args) {
		CustomerExcelFile customerExcelFile = new CustomerExcelFile();
		customerExcelFile.generateExcelFile();
	}
	@Autowired
	private CustomerService customerService;
	
	public void generateExcelFile() {
		try {
			//create xlsx file
			List<Customer> customers = customerService.findAll();
			System.out.println(customers);
			
			Workbook workbook = new XSSFWorkbook();
			//create sheet
			Sheet sheet = workbook.createSheet("CustomerData");
			//create top row with column heading
			String [] columnHeadings= {"Customer Id","Customer Code","Customer Name","Customer Email","Customer Phone",
					"Customer Address","Customer Purchasers","BillingAddress","DefaultBillingAddressIndex",
					"ShippingAddress","DefaultShippingAddressIndex","Active","CreatedAt","UpdatedAt"};
			//make heading bold with color blue
			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setFontHeightInPoints((short) 12);
			headerFont.setColor(IndexedColors.DARK_BLUE.index);
			//create cell style with a bond
			CellStyle headerStyle = workbook.createCellStyle();
			headerStyle.setFont(headerFont);
			headerStyle.setFillPattern(FillPatternType.ALT_BARS);
			headerStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.index);
			//create header Row
			Row headerRow = sheet.createRow(0);
			//Iterate over columns heading to create columns
			for(int i=0;i<columnHeadings.length;i++) {
				Cell cell = headerRow.createCell(i);
				cell.setCellValue(columnHeadings[i]);
				cell.setCellStyle(headerStyle);
			}
			//Fill Data
			int rownum=1;
			for(Customer i: customers) {
				Row row = sheet.createRow(rownum++);
				row.createCell(0).setCellValue(i.getId());
				row.createCell(1).setCellValue(i.getCode());
				row.createCell(2).setCellValue(i.getName());
				row.createCell(3).setCellValue(i.getEmail());
				row.createCell(4).setCellValue(i.getPhone());
				row.createCell(5).setCellValue(i.getAddress().getAddress());
				row.createCell(6).setCellValue(i.getPurchasers().toString());
//				row.createCell(7).setCellValue(i.getBillingAddress().toString());
				row.createCell(8).setCellValue(i.getDefaultBillingAddressIndex());
//				row.createCell(9).setCellValue(i.getShippingAddress().toString());
				row.createCell(10).setCellValue(i.getDefaultShippingAddressIndex());
				row.createCell(11).setCellValue(i.getActive());
				row.createCell(12).setCellValue(i.getCreatedAt());
				row.createCell(13).setCellValue(i.getUpdatedAt());
			}
			//Auto size columns
			for(int i=0; i<columnHeadings.length;i++) {
				sheet.autoSizeColumn(i);
			}
			Sheet sheet2 = workbook.createSheet("Second");
			//write the output to the file
			FileOutputStream fileOut = new FileOutputStream("C:\\Users\\abhis\\Desktop\\Spring Java\\customer_excel_files\\CustomerDataNewDevice.xlsx");
			workbook.write(fileOut);
			fileOut.close();
			workbook.close();
			System.out.println("Completed");
		}
		catch (Exception e) {
			e.printStackTrace();
		}
	}

}

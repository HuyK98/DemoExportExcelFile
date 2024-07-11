package mks.platform.samplewebapp.controller;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;
import mks.platform.samplewebapp.common.model.TableStructure;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpSession;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


@Controller
public class HansontableController extends BaseController {
	
	private String[] productListColHeaders = {"Account", "Section", "Mon", "Tue", "Wed", "Thur", "Fri", "Sat", "Sun"};
	

	private int[] productListColWidths = {100, 100, 100, 100, 100, 100, 100, 100, 100};

	private List<Object[]> lstProducts = new ArrayList<>();

	public HansontableController() {
		// Initialize with demo data
		lstProducts.add(new Object[] {"", "", "", "", "", "", "", "", ""});
	}

	@GetMapping(value = "/handsontable")
	public ModelAndView displayHome(HttpServletRequest request, HttpSession httpSession) {
		ModelAndView mav = new ModelAndView("handsontable");

		return mav;
	}
	
	@GetMapping(value = {"/handsontable/loaddata"}, produces="application/json")
	@ResponseBody
	public TableStructure getProductTableData() {
		TableStructure productTable = new TableStructure(productListColWidths, productListColHeaders, lstProducts);
		return productTable;
	}

	@PostMapping(value = "/handsontable/savedata", consumes = "application/json", produces = "application/json")
	@ResponseBody
	public String saveProductTableData(@RequestBody List<Object[]> newProductData) {
		if (newProductData != null && !newProductData.isEmpty()) {
			lstProducts.clear();
			lstProducts.addAll(newProductData);
			return "Data saved successfully";
		}
		return "No data to save";
	}
	@GetMapping("/handsontable/export/excel")
	public ResponseEntity<byte[]> exportToExcel() throws IOException {
	    Workbook workbook = new XSSFWorkbook();
	    Sheet sheet = workbook.createSheet("Handsontable Data");

	    // Create header row
	    Row headerRow = sheet.createRow(0);
	    for (int i = 0; i < productListColHeaders.length; i++) {
	        Cell cell = headerRow.createCell(i);
	        cell.setCellValue(productListColHeaders[i]);
	    }

	    // Fill data rows
	    for (int i = 0; i < lstProducts.size(); i++) {
	        Row dataRow = sheet.createRow(i + 1);
	        Object[] data = lstProducts.get(i);
	        for (int j = 0; j < data.length; j++) {
	            Cell cell = dataRow.createCell(j);
	            cell.setCellValue(data[j] != null ? data[j].toString() : "");
	        }
	    }

	    // Write the output to a byte array
	    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
	    workbook.write(outputStream);
	    workbook.close();

	    // Create the response
	    HttpHeaders headers = new HttpHeaders();
	    headers.add("Content-Disposition", "attachment; filename=handsontable-data.xlsx");
	    return new ResponseEntity<>(outputStream.toByteArray(), headers, HttpStatus.OK);
	}
}

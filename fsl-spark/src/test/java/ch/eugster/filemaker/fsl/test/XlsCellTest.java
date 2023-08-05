package ch.eugster.filemaker.fsl.test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeoutException;

import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaParsingWorkbook;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaRenderingWorkbook;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.eval.FunctionEval;
import org.apache.poi.ss.formula.ptg.AreaPtgBase;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.formula.ptg.RefPtgBase;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.jetty.client.api.ContentResponse;
import org.eclipse.jetty.client.util.StringContentProvider;
import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.fasterxml.jackson.databind.node.TextNode;

import ch.eugster.filemaker.fsl.Executor;
import ch.eugster.filemaker.fsl.xls.Key;
import ch.eugster.filemaker.fsl.xls.Xls;

public final class XlsCellTest extends AbstractXlsTest
{
	@Test
	public void testSetCellsWithoutWorkbook() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook  = null;
		
		ObjectNode requestNode = mapper.createObjectNode();
		ObjectNode cellNode = requestNode.objectNode();
		cellNode.put("row", 1);
		cellNode.put("col", 1);
		requestNode.set("cell", cellNode);
		TextNode directionNode = requestNode.textNode(Key.RIGHT.key());
		requestNode.set("direction", directionNode);

		ArrayNode valuesNode = requestNode.arrayNode();
		valuesNode.add("Title");
		valuesNode.add(1);
		valuesNode.add(2);
		valuesNode.add(3);
		valuesNode.add(4);
		valuesNode.add("SUM(C2:F2)");
		requestNode.set("values", valuesNode);
		requestNode.put("path", "./src/test/results/test.xlsx");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.setCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("workbook missing (create workbook first)", responseNode.get(Executor.ERRORS).get(0).asText());
	}

	@Test
	public void testSetCellsWithoutSheet() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook  = new XSSFWorkbook();

		ObjectNode requestNode = mapper.createObjectNode();
		ObjectNode cellNode = requestNode.objectNode();
		cellNode.put("row", 1);
		cellNode.put("col", 1);
		requestNode.set("cell", cellNode);
		TextNode directionNode = requestNode.textNode(Key.RIGHT.key());
		requestNode.set("direction", directionNode);

		ArrayNode valuesNode = requestNode.arrayNode();
		valuesNode.add("Title");
		valuesNode.add(1);
		valuesNode.add(2);
		valuesNode.add(3);
		valuesNode.add(4);
		valuesNode.add("SUM(C2:F2)");
		requestNode.set("values", valuesNode);
		requestNode.put("path", "./src/test/results/test.xlsx");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.setCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		ArrayNode errorsNode = ArrayNode.class.cast(responseNode.get(Executor.ERRORS));
		assertEquals(2, errorsNode.size());
		Iterator<JsonNode> iterator = errorsNode.iterator();
		while (iterator.hasNext())
		{
			JsonNode errorNode = iterator.next();
			if ("sheet index (0) is out of range (no sheets)".equals(errorNode.asText()))
			{
				assertTrue(true);
			}
			else if ("missing_argument 'sheet'".equals(errorNode.asText()))
			{
				assertTrue(true);
			}
			else
			{
				assertTrue(false);
			}
		}
	}

	@Test
	public void testSetCellsRightByIntegers() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook  = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		
		ObjectNode requestNode = mapper.createObjectNode();
		ObjectNode cellNode = requestNode.objectNode();
		cellNode.put("row", 1);
		cellNode.put("col", 1);
		requestNode.set("cell", cellNode);
		TextNode directionNode = requestNode.textNode(Key.RIGHT.key());
		requestNode.set("direction", directionNode);

		ArrayNode valuesNode = requestNode.arrayNode();
		valuesNode.add("Title");
		valuesNode.add(1);
		valuesNode.add(2);
		valuesNode.add(3);
		valuesNode.add(4);
		valuesNode.add("SUM(C2:F2)");
		requestNode.set("values", valuesNode);
		requestNode.put("path", "./src/test/results/test.xlsx");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.setCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(10, sheet.getRow(1).getCell(6).getNumericCellValue(), 0);
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testSetCellsRight() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook  = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		
		CellAddress cellAddress = new CellAddress("B2");
		ObjectNode requestNode = mapper.createObjectNode();
		TextNode startNode = requestNode.textNode(cellAddress.formatAsString());
		requestNode.set("cell", startNode);
		TextNode directionNode = requestNode.textNode(Key.RIGHT.key());
		requestNode.set("direction", directionNode);

		ArrayNode valuesNode = requestNode.arrayNode();
		valuesNode.add("Title");
		valuesNode.add(1);
		valuesNode.add(2);
		valuesNode.add(3);
		valuesNode.add(4);
		valuesNode.add("SUM(C2:F2)");
		requestNode.set("values", valuesNode);
		requestNode.put("path", "./src/test/results/test.xlsx");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.setCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(10, sheet.getRow(1).getCell(6).getNumericCellValue(), 0);
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testSetCellsRightOneCell() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook  = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		
		CellAddress cellAddress = new CellAddress("B2");
		ObjectNode requestNode = mapper.createObjectNode();
		TextNode firstNode = requestNode.textNode(cellAddress.formatAsString());
		requestNode.set("cell", firstNode);

		ArrayNode valuesNode = requestNode.arrayNode();
		valuesNode.add("Title");
		requestNode.set("values", valuesNode);
		requestNode.put("path", "./src/test/results/CellsRightOneCell.xlsx");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.setCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals("Title", sheet.getRow(1).getCell(1).getStringCellValue());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testSetCellsLeft() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook  = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		
		CellAddress cellAddress = new CellAddress("G3");
		ObjectNode requestNode = mapper.createObjectNode();
		TextNode startNode = requestNode.textNode(cellAddress.formatAsString());
		requestNode.set("cell", startNode);
		TextNode directionNode = requestNode.textNode(Key.LEFT.key());
		requestNode.set("direction", directionNode);

		ArrayNode valuesNode = requestNode.arrayNode();
		valuesNode.add("Title");
		valuesNode.add(1);
		valuesNode.add(2);
		valuesNode.add(3);
		valuesNode.add(4);
		valuesNode.add("SUM(C3:F3)");
		requestNode.set("values", valuesNode);
		requestNode.put("path", "./src/test/results/test.xlsx");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.setCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(10, sheet.getRow(2).getCell(1).getNumericCellValue(), 0);
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testSetCellsUp() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook  = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		
		CellAddress cellAddress = new CellAddress("I30");
		ObjectNode requestNode = mapper.createObjectNode();
		TextNode startNode = requestNode.textNode(cellAddress.formatAsString());
		requestNode.set("cell", startNode);
		TextNode directionNode = requestNode.textNode("up");
		requestNode.set("direction", directionNode);

		ArrayNode valuesNode = requestNode.arrayNode();
		valuesNode.add("Title");
		valuesNode.add(1);
		valuesNode.add(2);
		valuesNode.add(3);
		valuesNode.add(4);
		valuesNode.add("SUM(I26:I29)");
		requestNode.set("values", valuesNode);
		requestNode.put("path", "./src/test/results/test.xlsx");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.setCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(10, sheet.getRow(24).getCell(8).getNumericCellValue(), 0);
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testSetCellsDown() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook  = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		
		CellAddress cellAddress = new CellAddress("K3");
		ObjectNode requestNode = mapper.createObjectNode();
		TextNode startNode = requestNode.textNode(cellAddress.formatAsString());
		requestNode.set("cell", startNode);
		TextNode directionNode = requestNode.textNode("down");
		requestNode.set("direction", directionNode);

		ArrayNode valuesNode = requestNode.arrayNode();
		valuesNode.add("Title");
		valuesNode.add(1);
		valuesNode.add(2);
		valuesNode.add(3);
		valuesNode.add(4);
		valuesNode.add("SUM(K4:K7)");
		requestNode.set("values", valuesNode);
		requestNode.put("path", "./src/test/results/test.xlsx");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.setCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(10, sheet.getRow(7).getCell(10).getNumericCellValue(), 0);
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testSetCellsWithAddressValues() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook  = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		
		CellAddress cellAddress = new CellAddress("K3");
		ObjectNode requestNode = mapper.createObjectNode();
		ObjectNode cellNode = requestNode.objectNode();
		cellNode.put("row", cellAddress.getRow());
		cellNode.put("col", cellAddress.getColumn());
		requestNode.set("cell", cellNode);
		TextNode directionNode = requestNode.textNode("down");
		requestNode.set("direction", directionNode);

		ArrayNode valuesNode = requestNode.arrayNode();
		valuesNode.add("Title");
		valuesNode.add(1);
		valuesNode.add(2);
		valuesNode.add(3);
		valuesNode.add(4);
		valuesNode.add("SUM(K4:K7)");
		requestNode.set("values", valuesNode);
		requestNode.put("path", "./src/test/results/test.xlsx");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.setCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(10, sheet.getRow(7).getCell(10).getNumericCellValue(), 0);
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testSetCellsWithoutDirection() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook  = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		
		CellAddress cellAddress = new CellAddress("B2");
		ObjectNode requestNode = mapper.createObjectNode();
		TextNode startNode = requestNode.textNode(cellAddress.formatAsString());
		requestNode.set("cell", startNode);

		ArrayNode valuesNode = requestNode.arrayNode();
		valuesNode.add("Title");
		valuesNode.add(1);
		valuesNode.add(2);
		valuesNode.add(3);
		valuesNode.add(4);
		valuesNode.add("SUM(C2:F2)");
		requestNode.set("values", valuesNode);
		requestNode.put("workbook", WORKBOOK_1);

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.setCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(10, sheet.getRow(1).getCell(6).getNumericCellValue(), 0);
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testCopyFormula() throws InterruptedException, TimeoutException, ExecutionException, JsonMappingException, JsonProcessingException
	{
		Xls.activeWorkbook  = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		
		Row row = sheet.createRow(2);
		Cell cell = row.createCell(2);
		cell.setCellValue(1D);
		cell = row.createCell(3);
		cell.setCellValue(2D);
		cell = row.createCell(4);
		cell.setCellFormula("C3/D3");

		row = sheet.createRow(3);
		cell = row.createCell(2);
		cell.setCellValue(3D);
		cell = row.createCell(3);
		cell.setCellValue(4D);
		cell = row.createCell(4);
		cell.setCellFormula("$C4/D$4");

		row = sheet.createRow(4);
		cell = row.createCell(2);
		cell.setCellValue(5D);
		cell = row.createCell(3);
		cell.setCellValue(6D);
		cell = row.createCell(4);
		cell.setCellFormula("SUM(C3:D5)");

		row = sheet.createRow(5);
		cell = row.createCell(2);
		cell.setCellValue(7D);
		cell = row.createCell(3);
		cell.setCellValue(8D);
		cell = row.createCell(4);
		cell.setCellFormula("SUM(C$3/$D6)");

		row = sheet.createRow(6);
		cell = row.createCell(2);
		cell.setCellValue(9D);
		cell = row.createCell(3);
		cell.setCellValue(10D);
		cell = row.createCell(4);
		cell.setCellFormula("C3+SUM(C3:D7)");

		row = sheet.createRow(7);
		cell = row.createCell(2);
		cell.setCellValue(11D);
		cell = row.createCell(3);
		cell.setCellValue(12D);
		cell = row.createCell(4);
		cell.setCellFormula("C$3+SUM($C3:D$8)");

		for (Row r : sheet)
		{
			for (Cell c : r)
			{
				if (c.getCellType() == CellType.FORMULA)
				{
					CellAddress source = c.getAddress();
					String formula = c.getCellFormula();
					System.out.print(source + "=" + formula);
					int rowdiff = 3;
					int coldiff = -2;
					CellAddress target = new CellAddress(source.getRow() + rowdiff, source.getColumn() + coldiff);
					String newformula = copyFormula(sheet, formula, coldiff, rowdiff);
					System.out.println("->" + target + "=" + newformula);
				}
			}
		}
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), WORKBOOK_1);

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.saveWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		mapper.readTree(response.getContentAsString());
	}
	
	@Test
	public void testTimeAsStringToCell() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook = new XSSFWorkbook();
		Xls.activeWorkbook .createSheet();
		
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put("cell", "A1");
		ArrayNode valuesNode = requestNode.arrayNode();
		valuesNode.add("12:33");
		requestNode.set("values", valuesNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.setCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testDateAsStringToCell() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook  = new XSSFWorkbook();
		Xls.activeWorkbook .createSheet();
		
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put("cell", "A1");
		ArrayNode valuesNode = requestNode.arrayNode();
		valuesNode.add("21.10.1954");
		requestNode.set("values", valuesNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.setCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testDateTimeAsStringToCell() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook  = new XSSFWorkbook();
		Xls.activeWorkbook .createSheet();
		
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put("cell", "A1");
		ArrayNode valuesNode = requestNode.arrayNode();
		valuesNode.add("21.10.1954 10:31");
		requestNode.set("values", valuesNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.setCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testCopySingleFormulaCellToSingleCell() throws IOException, InterruptedException, TimeoutException, ExecutionException
	{
		String workbook = "./src/test/results/CopySingleFormulaCell.xlsx";
		Xls.activeWorkbook  = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		Row row0 = sheet.createRow(0);
		Cell cell = row0.createCell(0);
		cell.setCellValue(23.5);
		Row row1 = sheet.createRow(1);
		cell = row1.createCell(0);
		cell.setCellValue(76.5);
		Row row2 = sheet.createRow(2);
		cell = row2.createCell(0);
		cell.setCellFormula("SUM(A1:A2)");
		cell = row0.createCell(1);
		cell.setCellValue(12.5);
		cell = row1.createCell(1);
		cell.setCellValue(12.5);

		ObjectNode requestNode = mapper.createObjectNode();
		TextNode sourceNode = requestNode.textNode("A3");
		requestNode.set(Key.SOURCE.key(), sourceNode);

		TextNode targetNode = requestNode.textNode("B3");
		requestNode.set(Key.TARGET.key(), targetNode);

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.copyCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		cell = sheet.getRow(2).getCell(1);
		FormulaEvaluator formulaEval = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		assertEquals(25D, formulaEval.evaluate(cell).getNumberValue(), 0D);

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), workbook);

		response = client.POST("http://localhost:4567/fsl/Xls.saveWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertTrue(new File(workbook).isFile());
	}

	@Test
	public void testCopySingleFormulaCellToMultipleCells() throws IOException, InterruptedException, TimeoutException, ExecutionException
	{
		String path = "src/test/results/CopySingleCellMultipleCells.xlsx";
		Xls.activeWorkbook  = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		Row row0 = sheet.createRow(0);
		Cell cell = row0.createCell(0);
		cell.setCellValue(23.5);
		Row row1 = sheet.createRow(1);
		cell = row1.createCell(0);
		cell.setCellValue(76.5);
		Row row2 = sheet.createRow(2);
		cell = row2.createCell(0);
		cell.setCellFormula("SUM(A1:A2)");
		cell = row0.createCell(1);
		cell.setCellValue(12.5);
		cell = row1.createCell(1);
		cell.setCellValue(12.5);
		cell = row1.createCell(2);
		cell.setCellFormula("A3");
		cell = row1.createCell(2);
		cell.setCellFormula("B3");

		ObjectNode requestNode = mapper.createObjectNode();
		TextNode sourceNode = requestNode.textNode("A3");
		requestNode.set(Key.SOURCE.key(), sourceNode);

		TextNode targetNode = requestNode.textNode("B3:C3");
		requestNode.set(Key.TARGET.key(), targetNode);

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.copyCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		FormulaEvaluator formulaEval = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		cell = sheet.getRow(2).getCell(1);
		assertEquals(25D, formulaEval.evaluate(cell).getNumberValue(), 0D);
		cell = sheet.getRow(2).getCell(2);
		assertEquals(25D, formulaEval.evaluate(cell).getNumberValue(), 0D);

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), path);

		response = client.POST("http://localhost:4567/fsl/Xls.saveWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertTrue(new File(path).isFile());
	}

	@Test
	public void testCopySingleFormulaCellToMultipleCellsWithAddresses() throws IOException, InterruptedException, TimeoutException, ExecutionException
	{
		String path = "./src/test/results/testCopySingleFormulaCellToMultipleCellsWithAddresses.xlsx";
		Xls.activeWorkbook  = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		Row row0 = sheet.createRow(0);
		Cell cell = row0.createCell(0);
		cell.setCellValue(23.5);
		Row row1 = sheet.createRow(1);
		cell = row1.createCell(0);
		cell.setCellValue(76.5);
		Row row2 = sheet.createRow(2);
		cell = row2.createCell(0);
		cell.setCellFormula("SUM(A1:A2)");
		cell = row0.createCell(1);
		cell.setCellValue(12.5);
		cell = row1.createCell(1);
		cell.setCellValue(12.5);
		cell = row0.createCell(2);
		cell.setCellFormula("A3");
		cell = row1.createCell(2);
		cell.setCellFormula("B3");

		ObjectNode requestNode = mapper.createObjectNode();
		ObjectNode sourceNode = requestNode.objectNode();
		sourceNode.put(Key.TOP_LEFT.key(), "A3");
		sourceNode.put(Key.BOTTOM_RIGHT.key(), "A3");
		requestNode.set(Key.SOURCE.key(), sourceNode);

		ObjectNode targetNode = requestNode.objectNode();
		targetNode.put("top_left", "B3");
		targetNode.put(Key.BOTTOM_RIGHT.key(), "C3");
		requestNode.set(Key.TARGET.key(), targetNode);

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.copyCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		FormulaEvaluator formulaEval = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		cell = sheet.getRow(2).getCell(1);
		assertEquals(25D, formulaEval.evaluate(cell).getNumberValue(), 0D);
		cell = sheet.getRow(2).getCell(2);
		assertEquals(125D, formulaEval.evaluate(cell).getNumberValue(), 0D);

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), path);
	
		response = client.POST("http://localhost:4567/fsl/Xls.saveWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertTrue(new File(path).isFile());
	}

	@Test
	public void testCopySingleFormulaCellToMultipleCellsWithInts() throws IOException, InterruptedException, TimeoutException, ExecutionException
	{
		String path = "./src/test/results/testCopySingleFormulaCellToMultipleCellsWithInts.xlsx";
		Xls.activeWorkbook  = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		Row row0 = sheet.createRow(0);
		Cell cell = row0.createCell(0);
		cell.setCellValue(23.5);
		Row row1 = sheet.createRow(1);
		cell = row1.createCell(0);
		cell.setCellValue(76.5);
		Row row2 = sheet.createRow(2);
		cell = row2.createCell(0);
		cell.setCellFormula("SUM(A1:A2)");
		cell = row0.createCell(1);
		cell.setCellValue(12.5);
		cell = row1.createCell(1);
		cell.setCellValue(12.5);
		cell = row2.createCell(1);
		cell.setCellFormula("SUM(B1:B2)");
		cell = row0.createCell(2);
		cell.setCellFormula("A3");
		cell = row1.createCell(2);
		cell.setCellFormula("B3");
		cell = row2.createCell(2);
		cell.setCellFormula("SUM(C1:C2)");
		
		ObjectNode requestNode = mapper.createObjectNode();
		ObjectNode sourceNode = requestNode.objectNode();
		sourceNode.put(Key.TOP.key(), 2);
		sourceNode.put(Key.LEFT.key(), 0);
		sourceNode.put(Key.BOTTOM.key(), 2);
		sourceNode.put(Key.RIGHT.key(), 0);
		requestNode.set(Key.SOURCE.key(), sourceNode);

		ObjectNode targetNode = requestNode.objectNode();
		targetNode.put(Key.TOP.key(), 2);
		targetNode.put(Key.LEFT.key(), 1);
		targetNode.put(Key.BOTTOM.key(), 2);
		targetNode.put(Key.RIGHT.key(), 2);
		requestNode.set(Key.TARGET.key(), targetNode);

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.copyCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		FormulaEvaluator formulaEval = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		cell = sheet.getRow(2).getCell(1);
		assertEquals(25D, formulaEval.evaluate(cell).getNumberValue(), 0D);
		cell = sheet.getRow(2).getCell(2);
//		assertEquals(125D, formulaEval.evaluate(cell).getNumberValue(), 0D);

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), path);

		response = client.POST("http://localhost:4567/fsl/Xls.saveWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertTrue(new File(path).isFile());
	}

	@Test
	public void testSupportedFormulaNames()
	{
		for (String supportedFormulaName : FunctionEval.getNotSupportedFunctionNames())
		{
			System.out.println(supportedFormulaName);
		}
	}

	@Test
	public void testFunctionSupported()
	{
		String function = "SUM(A1:B2)";
		int pos = function.indexOf("(");
		if (pos > -1)
		{
			String name = function.substring(0, pos);
			for (String supportedFormulaName : FunctionEval.getSupportedFunctionNames())
			{
				if (supportedFormulaName.equals(name))
				{
					assertTrue(true);
					return;
				}
			}
		}
		assertFalse(true);
	}

	protected String copyFormula(Sheet sheet, String formula, int rowDiff, int colDiff)
	{
		FormulaParsingWorkbook workbookWrapper = getFormulaParsingWorkbook(sheet);
		Ptg[] ptgs = FormulaParser.parse(formula, workbookWrapper, FormulaType.CELL,
				sheet.getWorkbook().getSheetIndex(sheet));
		for (int i = 0; i < ptgs.length; i++)
		{
			if (ptgs[i] instanceof RefPtgBase)
			{ // base class for cell references
				RefPtgBase ref = (RefPtgBase) ptgs[i];
				if (ref.isRowRelative())
				{
					ref.setRow(ref.getRow() + rowDiff);
				}
				if (ref.isColRelative())
				{
					ref.setColumn(ref.getColumn() + colDiff);
				}
			}
			else if (ptgs[i] instanceof AreaPtgBase)
			{ // base class for range references
				AreaPtgBase ref = (AreaPtgBase) ptgs[i];
				if (ref.isFirstColRelative())
				{
					ref.setFirstColumn(ref.getFirstColumn() + colDiff);
				}
				if (ref.isLastColRelative())
				{
					ref.setLastColumn(ref.getLastColumn() + colDiff);
				}
				if (ref.isFirstRowRelative())
				{
					ref.setFirstRow(ref.getFirstRow() + rowDiff);
				}
				if (ref.isLastRowRelative())
				{
					ref.setLastRow(ref.getLastRow() + rowDiff);
				}
			}
		}

		formula = FormulaRenderer.toFormulaString(getFormulaRenderingWorkbook(sheet), ptgs);
		return formula;
	}

	protected FormulaParsingWorkbook getFormulaParsingWorkbook(Sheet sheet)
	{
		FormulaParsingWorkbook workbookWrapper = null;
		if (XSSFSheet.class.isInstance(sheet))
		{
			workbookWrapper = XSSFEvaluationWorkbook.create(XSSFSheet.class.cast(sheet).getWorkbook());
		}
		else
		{
			workbookWrapper = HSSFEvaluationWorkbook.create(HSSFSheet.class.cast(sheet).getWorkbook());
		}
		return workbookWrapper;
	}

	protected FormulaRenderingWorkbook getFormulaRenderingWorkbook(Sheet sheet)
	{
		FormulaRenderingWorkbook workbookWrapper = null;
		if (XSSFSheet.class.isInstance(sheet))
		{
			workbookWrapper = XSSFEvaluationWorkbook.create(XSSFSheet.class.cast(sheet).getWorkbook());
		}
		else
		{
			workbookWrapper = HSSFEvaluationWorkbook.create(HSSFSheet.class.cast(sheet).getWorkbook());
		}
		return workbookWrapper;
	}

}

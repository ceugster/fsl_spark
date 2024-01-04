package ch.eugster.filemaker.fsl.test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

import java.io.IOException;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeoutException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.jetty.client.api.ContentResponse;
import org.eclipse.jetty.client.util.StringContentProvider;
import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.fasterxml.jackson.databind.node.TextNode;

import ch.eugster.filemaker.fsl.Executor;
import ch.eugster.filemaker.fsl.xls.Key;
import ch.eugster.filemaker.fsl.xls.Xls;

public final class XlsCellStyleTest extends AbstractXlsTest
{
	@Test
	public void testHorizontalAlignment() throws EncryptedDocumentException, IOException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook.createSheet();
		Row row = sheet.createRow(1);
		Cell leftCell = row.createCell(1);
		leftCell.setCellValue(new XSSFRichTextString("Linksausrichtung"));
		Cell centerCell = row.createCell(3);
		centerCell.setCellValue(new XSSFRichTextString("Mitteausrichtung"));
		Cell rightCell = row.createCell(5);
		rightCell.setCellValue(new XSSFRichTextString("Rechtsausrichtung"));
		Cell distributedCell = row.createCell(7);
		distributedCell.setCellValue(new XSSFRichTextString("Verteilt auf die ganze Breite"));
		Cell errorCell = row.createCell(9);
		errorCell.setCellValue(new XSSFRichTextString("Fehler"));
		
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "B2");
		ObjectNode alignNode = requestNode.objectNode();
		alignNode.put(Key.HORIZONTAL.key(), HorizontalAlignment.LEFT.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "D2");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.HORIZONTAL.key(), HorizontalAlignment.CENTER.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "F2");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.HORIZONTAL.key(), HorizontalAlignment.RIGHT.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "H2");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.HORIZONTAL.key(), HorizontalAlignment.DISTRIBUTED.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "J2");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.HORIZONTAL.key(), "gigi");
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("illegal argument 'alignment.horizontal' (gigi)", responseNode.get(Executor.ERRORS).get(0).asText());

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), "src/test/results/testHorizontalAlignment");
		
		response = client.POST("http://localhost:4567/fsl/Xls.saveAndReleaseWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}
	
	@Test
	public void testHorizontalAlignmentRange() throws EncryptedDocumentException, IOException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook.createSheet();
		Row row = sheet.createRow(1);
		Cell leftCell = row.createCell(1);
		leftCell.setCellValue(new XSSFRichTextString("Linksausrichtung"));
		Cell centerCell = row.createCell(3);
		centerCell.setCellValue(new XSSFRichTextString("Mitteausrichtung"));
		Cell rightCell = row.createCell(5);
		rightCell.setCellValue(new XSSFRichTextString("Rechtsausrichtung"));
		Cell distributedCell = row.createCell(7);
		distributedCell.setCellValue(new XSSFRichTextString("Verteilt auf die ganze Breite"));
		Cell errorCell = row.createCell(9);
		errorCell.setCellValue(new XSSFRichTextString("Fehler"));
		row = sheet.createRow(3);
		leftCell = row.createCell(1);
		leftCell.setCellValue(new XSSFRichTextString("Linksausrichtung"));
		centerCell = row.createCell(3);
		centerCell.setCellValue(new XSSFRichTextString("Mitteausrichtung"));
		rightCell = row.createCell(5);
		rightCell.setCellValue(new XSSFRichTextString("Rechtsausrichtung"));
		distributedCell = row.createCell(7);
		distributedCell.setCellValue(new XSSFRichTextString("Verteilt auf die ganze Breite"));
		errorCell = row.createCell(9);
		errorCell.setCellValue(new XSSFRichTextString("Fehler"));
		
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "B2:B4");
		ObjectNode alignNode = requestNode.objectNode();
		alignNode.put(Key.HORIZONTAL.key(), HorizontalAlignment.LEFT.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "D2:D4");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.HORIZONTAL.key(), HorizontalAlignment.CENTER.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "F2:F4");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.HORIZONTAL.key(), HorizontalAlignment.RIGHT.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "H2:H4");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.HORIZONTAL.key(), HorizontalAlignment.DISTRIBUTED.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "J2:J4");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.HORIZONTAL.key(), "gigi");
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("illegal argument 'alignment.horizontal' (gigi)", responseNode.get(Executor.ERRORS).get(0).asText());

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), "src/test/results/testHorizontalAlignmentRange");

		response = client.POST("http://localhost:4567/fsl/Xls.saveAndReleaseWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}
	
	@Test
	public void testVerticalAlignment() throws EncryptedDocumentException, IOException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		Row row = sheet.createRow(1);
		Cell leftCell = row.createCell(1);
		leftCell.setCellValue(new XSSFRichTextString("Top"));
		Cell centerCell = row.createCell(3);
		centerCell.setCellValue(new XSSFRichTextString("Center"));
		Cell rightCell = row.createCell(5);
		rightCell.setCellValue(new XSSFRichTextString("Bottom"));
		Cell distributedCell = row.createCell(7);
		distributedCell.setCellValue(new XSSFRichTextString("Bottom"));
		Cell errorCell = row.createCell(9);
		errorCell.setCellValue(new XSSFRichTextString("Fehler"));
		
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "B2");
		ObjectNode alignNode = requestNode.objectNode();
		alignNode.put(Key.VERTICAL.key(), VerticalAlignment.TOP.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "D2");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.VERTICAL.key(), VerticalAlignment.CENTER.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "F2");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.VERTICAL.key(), VerticalAlignment.BOTTOM.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "H2");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.VERTICAL.key(), VerticalAlignment.DISTRIBUTED.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "J2");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.VERTICAL.key(), "gigi");
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("illegal argument 'alignment.vertical' (gigi)", responseNode.get(Executor.ERRORS).get(0).asText());

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), "src/test/results/testVerticalAlignment");

		response = client.POST("http://localhost:4567/fsl/Xls.saveAndReleaseWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}
	
	@Test
	public void testVerticalAlignmentRange() throws EncryptedDocumentException, IOException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		Row row = sheet.createRow(1);
		Cell leftCell = row.createCell(1);
		leftCell.setCellValue(new XSSFRichTextString("Top"));
		Cell centerCell = row.createCell(3);
		centerCell.setCellValue(new XSSFRichTextString("Center"));
		Cell rightCell = row.createCell(5);
		rightCell.setCellValue(new XSSFRichTextString("Bottom"));
		Cell distributedCell = row.createCell(7);
		distributedCell.setCellValue(new XSSFRichTextString("Bottom"));
		Cell errorCell = row.createCell(9);
		errorCell.setCellValue(new XSSFRichTextString("Fehler"));
		row = sheet.createRow(3);
		leftCell = row.createCell(1);
		leftCell.setCellValue(new XSSFRichTextString("Top"));
		centerCell = row.createCell(3);
		centerCell.setCellValue(new XSSFRichTextString("Center"));
		rightCell = row.createCell(5);
		rightCell.setCellValue(new XSSFRichTextString("Bottom"));
		distributedCell = row.createCell(7);
		distributedCell.setCellValue(new XSSFRichTextString("Bottom"));
		errorCell = row.createCell(9);
		errorCell.setCellValue(new XSSFRichTextString("Fehler"));
		
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "B2:B4");
		ObjectNode alignNode = requestNode.objectNode();
		alignNode.put(Key.VERTICAL.key(), VerticalAlignment.TOP.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "D2:D4");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.VERTICAL.key(), VerticalAlignment.CENTER.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "F2:F4");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.VERTICAL.key(), VerticalAlignment.BOTTOM.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "H2:H4");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.VERTICAL.key(), VerticalAlignment.DISTRIBUTED.name().toLowerCase());
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "J2:J4");
		alignNode = requestNode.objectNode();
		alignNode.put(Key.VERTICAL.key(), "gigi");
		requestNode.set(Key.ALIGNMENT.key(), alignNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("illegal argument 'alignment.vertical' (gigi)", responseNode.get(Executor.ERRORS).get(0).asText());

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), "src/test/results/testVerticalAlignmentRange");

		response = client.POST("http://localhost:4567/fsl/Xls.saveAndReleaseWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}
	
	@Test
	public void testBorderStyles() throws EncryptedDocumentException, IOException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook = new XSSFWorkbook();
		Xls.activeWorkbook .createSheet();
		
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "B2");
		ObjectNode borderNode = requestNode.objectNode();
		ObjectNode styleNode = borderNode.objectNode();
		TextNode node = styleNode.textNode(BorderStyle.DOTTED.name());
		styleNode.set(Key.BOTTOM.key(), node);
		borderNode.set(Key.STYLE.key(), styleNode);
		requestNode.set(Key.BORDER.key(), borderNode);

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "D4");
		borderNode = requestNode.objectNode();
		styleNode = borderNode.objectNode();
		node = styleNode.textNode(BorderStyle.THICK.name());
		styleNode.set(Key.LEFT.key(), node);
		borderNode.set(Key.STYLE.key(), styleNode);
		requestNode.set(Key.BORDER.key(), borderNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "F6");
		borderNode = requestNode.objectNode();
		styleNode = borderNode.objectNode();
		node = styleNode.textNode(BorderStyle.DASH_DOT.name());
		styleNode.set(Key.RIGHT.key(), node);
		borderNode.set(Key.STYLE.key(), styleNode);
		requestNode.set(Key.BORDER.key(), borderNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "H8");
		borderNode = requestNode.objectNode();
		styleNode = borderNode.objectNode();
		node = styleNode.textNode(BorderStyle.DASH_DOT.name());
		styleNode.set(Key.RIGHT.key(), node);
		borderNode.set(Key.STYLE.key(), styleNode);
		requestNode.set(Key.BORDER.key(), borderNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "J10");
		borderNode = requestNode.objectNode();
		styleNode = borderNode.objectNode();
		node = styleNode.textNode("gigi");
		styleNode.set(Key.TOP.key(), node);
		borderNode.set(Key.STYLE.key(), styleNode);
		requestNode.set(Key.BORDER.key(), borderNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("illegal argument 'border.style.top' (gigi)", responseNode.get(Executor.ERRORS).get(0).asText());

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), "src/test/results/testBorderStyles");

		response = client.POST("http://localhost:4567/fsl/Xls.saveAndReleaseWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}
	
	@Test
	public void testBorderStylesRange() throws EncryptedDocumentException, IOException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook = new XSSFWorkbook();
		Xls.activeWorkbook .createSheet();

		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "B2:C3");
		ObjectNode borderNode = requestNode.objectNode();
		ObjectNode styleNode = borderNode.objectNode();
		TextNode node = styleNode.textNode(BorderStyle.DOTTED.name());
		styleNode.set(Key.BOTTOM.key(), node);
		borderNode.set(Key.STYLE.key(), styleNode);
		requestNode.set(Key.BORDER.key(), borderNode);

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "E5:F6");
		borderNode = requestNode.objectNode();
		styleNode = borderNode.objectNode();
		node = styleNode.textNode(BorderStyle.THICK.name());
		styleNode.set(Key.LEFT.key(), node);
		borderNode.set(Key.STYLE.key(), styleNode);
		requestNode.set(Key.BORDER.key(), borderNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "H8:I9");
		borderNode = requestNode.objectNode();
		styleNode = borderNode.objectNode();
		node = styleNode.textNode(BorderStyle.DASH_DOT.name());
		styleNode.set(Key.RIGHT.key(), node);
		borderNode.set(Key.STYLE.key(), styleNode);
		requestNode.set(Key.BORDER.key(), borderNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "K11:L12");
		borderNode = requestNode.objectNode();
		styleNode = borderNode.objectNode();
		node = styleNode.textNode(BorderStyle.DASH_DOT.name());
		styleNode.set(Key.RIGHT.key(), node);
		borderNode.set(Key.STYLE.key(), styleNode);
		requestNode.set(Key.BORDER.key(), borderNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "N14:O15");
		borderNode = requestNode.objectNode();
		styleNode = borderNode.objectNode();
		node = styleNode.textNode("gigi");
		styleNode.set(Key.TOP.key(), node);
		borderNode.set(Key.STYLE.key(), styleNode);
		requestNode.set(Key.BORDER.key(), borderNode);

		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("illegal argument 'border.style.top' (gigi)", responseNode.get(Executor.ERRORS).get(0).asText());

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), "src/test/results/testBorderStyleRange");

		response = client.POST("http://localhost:4567/fsl/Xls.saveAndReleaseWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testDataFormatNumber() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		String path = "./src/test/results/dataFormatNumber.xlsx";
		Xls.activeWorkbook = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue(12.666);
		cell = row.createCell(1);
		cell.setCellValue(13.6);
		cell = row.createCell(2);
		cell.setCellValue(1000.3);
		row = sheet.createRow(1);
		cell = row.createCell(0);
		cell.setCellValue(12.666);
		cell = row.createCell(1);
		cell.setCellValue(13.6);
		cell = row.createCell(2);
		cell.setCellValue(1000.3);
		row = sheet.createRow(2);
		cell = row.createCell(0);
		cell.setCellValue(12.666);
		cell = row.createCell(1);
		cell.setCellValue(13.6);
		cell = row.createCell(2);
		cell.setCellValue(1000.3);

		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "A1");
		requestNode.put(Key.DATA_FORMAT.key(), "0.00");
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "B1:C1");
		requestNode.put(Key.DATA_FORMAT.key(), "0.00");
		
		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "A2");
		requestNode.put(Key.DATA_FORMAT.key(), "#,##0.00");
		
		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "B2:C2");
		requestNode.put(Key.DATA_FORMAT.key(), "#,##0.00");
		
		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "A3");
		requestNode.put(Key.DATA_FORMAT.key(), "#,##0");
		
		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "B3:C3");
		requestNode.put(Key.DATA_FORMAT.key(), "#,##0");
		
		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), path);

		response = client.POST("http://localhost:4567/fsl/Xls.saveAndReleaseWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}
	
	@Test
	public void testDataFormatTime() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		String path = "./src/test/results/dataFormatTime.xlsx";
		Xls.activeWorkbook = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue("1:25");
		cell = row.createCell(1);
		cell.setCellValue("1:25:00");
		cell = row.createCell(2);
		cell.setCellValue("21.10.1954");
		cell = row.createCell(3);
		cell.setCellValue("21.10.1954 1:25");
		cell = row.createCell(4);
		cell.setCellValue(4.553472222);
		

		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "A1");
		requestNode.put(Key.DATA_FORMAT.key(), "h:mm");
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "B1");
		requestNode.put(Key.DATA_FORMAT.key(), "h:mm");
		
		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "E1");
		requestNode.put(Key.DATA_FORMAT.key(), "[h]:mm");
		
		response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), path);

		response = client.POST("http://localhost:4567/fsl/Xls.saveAndReleaseWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}
	
	@Test
	public void testBackground() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		String path = "src/test/results/background.xlsx";
		Xls.activeWorkbook = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue(new XSSFRichTextString("Das ist eine Testzelle"));
		cell = row.createCell(1);
		cell.setCellValue(new XSSFRichTextString("Das ist eine zweite Testzelle"));

		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "A1");
		ObjectNode backgroundNode = requestNode.objectNode();
		backgroundNode.put("color", IndexedColors.RED.name());
		requestNode.set(Key.BACKGROUND.key(), backgroundNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.applyCellStyles").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), path);

		response = client.POST("http://localhost:4567/fsl/Xls.saveAndReleaseWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}
	
	@Test
	public void testAutoSizeColumns() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		String path = "src/test/results/autoSizeColumn.xlsx";
		Xls.activeWorkbook = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue(new XSSFRichTextString("Das ist eine Testzelle"));
		cell = row.createCell(1);
		cell.setCellValue(new XSSFRichTextString("Das ist eine zweite Testzelle"));

		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Key.CELL.key(), "A1");
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.autoSizeColumns").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.RANGE.key(), "A1:B1");
		
		response = client.POST("http://localhost:4567/fsl/Xls.autoSizeColumns").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), path);

		response = client.POST("http://localhost:4567/fsl/Xls.saveAndReleaseWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
	}

	@Test
	public void testRotation() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		String path = "src/test/results/rotate.xlsx";
		Xls.activeWorkbook = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue(new XSSFRichTextString("Das ist eine Testzelle"));

		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Key.ROTATION.key(), 90);
		requestNode.put(Key.CELL.key(), "A1");
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.rotateCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), path);

		response = client.POST("http://localhost:4567/fsl/Xls.saveAndReleaseWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
	}
	
	@Test
	public void testRangeRotation() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		String path = "src/test/results/rotateRange.xlsx";
		Xls.activeWorkbook = new XSSFWorkbook();
		Sheet sheet = Xls.activeWorkbook .createSheet();
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue(new XSSFRichTextString("Das ist eine Testzelle"));
		cell = row.createCell(1);
		cell.setCellValue(new XSSFRichTextString("Das ist eine zweite Testzelle"));

		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Key.ROTATION.key(), 90);
		requestNode.put(Key.RANGE.key(), "A1:B1");
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.rotateCells").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.PATH.key(), path);

		response = client.POST("http://localhost:4567/fsl/Xls.saveAndReleaseWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
	}
	
}

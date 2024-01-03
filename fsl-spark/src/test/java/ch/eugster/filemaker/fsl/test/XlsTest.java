package ch.eugster.filemaker.fsl.test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Iterator;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeoutException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.jetty.client.api.ContentResponse;
import org.eclipse.jetty.client.util.StringContentProvider;
import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import ch.eugster.filemaker.fsl.Executor;
import ch.eugster.filemaker.fsl.xls.Key;
import ch.eugster.filemaker.fsl.xls.Xls;

public final class XlsTest extends AbstractXlsTest
{
	@Test
	public void testEmptyParameter() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.createWorkbook").
				header("Content-Type", "application/json").
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("missing or illegal argument", responseNode.get(Executor.ERRORS).get(0).asText());
	}
	
	@Test
	public void testSupportedFunctionNames() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.getSupportedFunctionNames").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals("OK", responseNode.get(Executor.STATUS).asText());
		assertEquals(185, responseNode.get(Executor.RESULT).size());
		assertNull(responseNode.get(Executor.ERRORS));
	}
	
	@Test
	public void testCallableMethods() throws Exception
	{
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.getCallableMethods").
				header("Content-Type", "application/json").
				content(new StringContentProvider(mapper.createObjectNode().toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		ArrayNode methods = ArrayNode.class.cast(responseNode.get(Executor.RESULT));
		for (int i = 0; i < methods.size(); i++)
		{
			System.out.println(methods.get(i).asText());
		}
	}
	
	@Test
	public void testWorkbookActive() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook = null;
		
		ObjectNode requestNode = mapper.createObjectNode();
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.workbookPresent").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(0, responseNode.get(Executor.RESULT).asInt());
		assertNull(responseNode.get(Executor.ERRORS));

		Xls.activeWorkbook = new XSSFWorkbook();
		
		response = client.POST("http://localhost:4567/fsl/Xls.workbookPresent").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.RESULT).asInt());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testActiveSheetPresent() throws Exception
	{
		Xls.activeWorkbook = null;
		
		ObjectNode requestNode = mapper.createObjectNode();
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.activeSheetPresent").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.RESULT));
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("workbook missing (create workbook first)", responseNode.get(Executor.ERRORS).get(0).asText());

		Xls.activeWorkbook = new XSSFWorkbook();

		response = client.POST("http://localhost:4567/fsl/Xls.activeSheetPresent").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(0, responseNode.get(Key.INDEX.key()).asInt());
		assertEquals("", responseNode.get(Key.SHEET.key()).asText());

		Xls.activeWorkbook.createSheet();
		
		response = client.POST("http://localhost:4567/fsl/Xls.activeSheetPresent").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Key.INDEX.key()).asInt());
		assertEquals("Sheet0", responseNode.get(Key.SHEET.key()).asText());
	}

	@Test
	public void testCreateWorkbook() throws Exception
	{
		Xls.activeWorkbook = null;
		
		ObjectNode requestNode = mapper.createObjectNode();

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.createWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		Xls.activeWorkbook = new XSSFWorkbook();
		
		response = client.POST("http://localhost:4567/fsl/Xls.createWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testCreateSheet() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook = null;
		
		ObjectNode requestNode = mapper.createObjectNode();

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.createSheet").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("workbook missing (create workbook first)", responseNode.get(Executor.ERRORS).get(0).asText());

		Xls.activeWorkbook = new XSSFWorkbook();

		response = client.POST("http://localhost:4567/fsl/Xls.createSheet").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(0, responseNode.get(Key.INDEX.key()).asInt());
		assertEquals(Xls.activeWorkbook.getSheetAt(0).getSheetName(), responseNode.get(Key.SHEET.key()).asText());

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.SHEET.key(), "Arbeitsblatt 1");
		
		response = client.POST("http://localhost:4567/fsl/Xls.createSheet").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Key.INDEX.key()).asInt());
		assertEquals(Xls.activeWorkbook.getSheetAt(1).getSheetName(), responseNode.get(Key.SHEET.key()).asText());

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.SHEET.key(), "Arbeitsblatt 1");
		
		response = client.POST("http://localhost:4567/fsl/Xls.createSheet").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Key.INDEX.key()));
		assertNull(responseNode.get(Key.SHEET.key()));
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("illegal argument 'sheet' ('Arbeitsblatt 1' already exists)", responseNode.get(Executor.ERRORS).get(0).asText());
	}

	@Test
	public void testDropSheet() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook = null;
		
		ObjectNode requestNode = mapper.createObjectNode();

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.dropSheet").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("workbook missing (create workbook first)", responseNode.get(Executor.ERRORS).get(0).asText());

		Xls.activeWorkbook = new XSSFWorkbook();

		response = client.POST("http://localhost:4567/fsl/Xls.dropSheet").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("there is no active sheet present", responseNode.get(Executor.ERRORS).get(0).asText());

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.SHEET.key(), "Arbeitsblatt 1");
		
		response = client.POST("http://localhost:4567/fsl/Xls.dropSheet").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("sheet with name 'Arbeitsblatt 1' does not exist", responseNode.get(Executor.ERRORS).get(0).asText());

		Xls.activeWorkbook.createSheet("Arbeitsblatt 1");
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.SHEET.key(), "Arbeitsblatt 1");
		
		response = client.POST("http://localhost:4567/fsl/Xls.dropSheet").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testGetSheetNames() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.sheetNames").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("workbook missing (create workbook first)", responseNode.get(Executor.ERRORS).get(0).asText());
		
		Xls.activeWorkbook = new XSSFWorkbook();
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.WORKBOOK.key(), WORKBOOK_1);

		response = client.POST("http://localhost:4567/fsl/Xls.sheetNames").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		JsonNode sheetNode = responseNode.get(Key.SHEET.key());
		assertTrue(ArrayNode.class.isInstance(sheetNode));
		assertEquals(0, sheetNode.size());
		JsonNode indexNode = responseNode.get("index");
		assertEquals(0, indexNode.size());
		assertNull(responseNode.get(Executor.ERRORS));

		Xls.activeWorkbook.createSheet("Arbeitsblatt 1");
		
		response = client.POST("http://localhost:4567/fsl/Xls.sheetNames").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		sheetNode = responseNode.get(Key.SHEET.key());
		assertTrue(ArrayNode.class.isInstance(sheetNode));
		assertEquals(1, sheetNode.size());
		Iterator<JsonNode> iterator = sheetNode.iterator();
		assertEquals("Arbeitsblatt 1", iterator.next().asText());
		indexNode = responseNode.get(Key.INDEX.key());
		assertEquals(1, indexNode.size());
		iterator = indexNode.iterator();
		assertEquals(0, iterator.next().asInt());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testActivateSheet() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		Xls.activeWorkbook = null;
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.activateSheet").
				header("Content-Type", "application/json").
				content(new StringContentProvider(mapper.createObjectNode().toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("workbook missing (create workbook first)", responseNode.get(Executor.ERRORS).get(0).asText());

		Xls.activeWorkbook = new XSSFWorkbook();
		
		response = client.POST("http://localhost:4567/fsl/Xls.activateSheet").
				header("Content-Type", "application/json").
				content(new StringContentProvider(mapper.createObjectNode().toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("missing argument 'sheet' or 'index'", responseNode.get(Executor.ERRORS).get(0).asText());

		Xls.activeWorkbook.createSheet();
		
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Key.SHEET.key(), "Arbeitsblatt 1");

		response = client.POST("http://localhost:4567/fsl/Xls.activateSheet").
				header("Content-Type", "application/json").
				content(new StringContentProvider(mapper.createObjectNode().toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("missing argument 'sheet' or 'index'", responseNode.get(Executor.ERRORS).get(0).asText());
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.SHEET.key(), SHEET0);

		response = client.POST("http://localhost:4567/fsl/Xls.activateSheet").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(0, responseNode.get(Key.INDEX.key()).asInt());
		assertEquals(SHEET0, responseNode.get(Key.SHEET.key()).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.INDEX.key(), 1);

		response = client.POST("http://localhost:4567/fsl/Xls.activateSheet").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("sheet with index 1 does not exist", responseNode.get(Executor.ERRORS).get(0).asText());
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.INDEX.key(), 0);

		response = client.POST("http://localhost:4567/fsl/Xls.activateSheet").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(0, responseNode.get(Key.INDEX.key()).asInt());
		assertEquals(SHEET0, responseNode.get(Key.SHEET.key()).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testSaveWorkbook() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.saveWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("workbook missing (create workbook first)", responseNode.get(Executor.ERRORS).get(0).asText());

		Xls.activeWorkbook = new XSSFWorkbook();
		
		response = client.POST("http://localhost:4567/fsl/Xls.saveWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("missing argument '" + Key.PATH.key() + "'",responseNode.get(Executor.ERRORS).get(0).asText());
	}

	@Test
	public void testReleaseWorkbook() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.releaseWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("workbook missing (create workbook first)",responseNode.get(Executor.ERRORS).get(0).asText());

		Xls.activeWorkbook = new XSSFWorkbook();
		
		response = client.POST("http://localhost:4567/fsl/Xls.releaseWorkbook").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testSetHeaders() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put("left", "Header links");
		requestNode.put("center", "Header Mitte");
		requestNode.put("right", "Header rechts");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.setHeaders").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("workbook missing (create workbook first)", responseNode.get(Executor.ERRORS).get(0).asText());

		Xls.activeWorkbook = new XSSFWorkbook();

		requestNode = mapper.createObjectNode();
		requestNode.put("left", "Header links");
		requestNode.put("center", "Header Mitte");
		requestNode.put("right", "Header rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setHeaders").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("sheet index (0) is out of range (no sheets)", responseNode.get(Executor.ERRORS).get(0).asText());
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.SHEET.key(), SHEET0);
		requestNode.put("left", "Header links");
		requestNode.put("center", "Header Mitte");
		requestNode.put("right", "Header rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setHeaders").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("sheet index (0) is out of range (no sheets)", responseNode.get(Executor.ERRORS).get(0).asText());
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.INDEX.key(), 0);
		requestNode.put("left", "Header links");
		requestNode.put("center", "Header Mitte");
		requestNode.put("right", "Header rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setHeaders").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("sheet index (0) is out of range (no sheets)", responseNode.get(Executor.ERRORS).get(0).asText());

		response = client.POST("http://localhost:4567/fsl/Xls.setHeaders").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		Xls.activeWorkbook.createSheet();
		
		responseNode = mapper.readTree(response.getContentAsString());
		requestNode.put("left", "Header links");
		requestNode.put("center", "Header Mitte");
		requestNode.put("right", "Header rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setHeaders").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.SHEET.key(), SHEET0);
		requestNode.put("left", "Header links");
		requestNode.put("center", "Header Mitte");
		requestNode.put("right", "Header rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setHeaders").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.INDEX.key(), 0);
		requestNode.put("left", "Header links");
		requestNode.put("center", "Header Mitte");
		requestNode.put("right", "Header rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setHeaders").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.SHEET.key(), "Arbeitsblatt 1");
		requestNode.put("left", "Header links");
		requestNode.put("center", "Header Mitte");
		requestNode.put("right", "Header rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setHeaders").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("sheet with name 'Arbeitsblatt 1' does not exist", responseNode.get(Executor.ERRORS).get(0).asText());
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.INDEX.key(), 1);
		requestNode.put("left", "Header links");
		requestNode.put("center", "Header Mitte");
		requestNode.put("right", "Header rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setHeaders").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("sheet index (1) is out of range (0..0)", responseNode.get(Executor.ERRORS).get(0).asText());

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.INDEX.key(), 0);
		requestNode.put("left", "Header links");
		requestNode.put("center", "Header Mitte");
		requestNode.put("right", "Header rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setHeaders").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testSetFooters() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put("left", "Footer links");
		requestNode.put("center", "Footer Mitte");
		requestNode.put("right", "Footer rechts");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.setFooters").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("workbook missing (create workbook first)", responseNode.get(Executor.ERRORS).get(0).asText());

		Xls.activeWorkbook = new XSSFWorkbook();

		requestNode = mapper.createObjectNode();
		requestNode.put("left", "Footer links");
		requestNode.put("center", "Footer Mitte");
		requestNode.put("right", "Footer rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setFooters").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("sheet index (0) is out of range (no sheets)", responseNode.get(Executor.ERRORS).get(0).asText());
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.SHEET.key(), SHEET0);
		requestNode.put("left", "Footer links");
		requestNode.put("center", "Footer Mitte");
		requestNode.put("right", "Footer rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setFooters").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("sheet index (0) is out of range (no sheets)", responseNode.get(Executor.ERRORS).get(0).asText());
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.INDEX.key(), 0);
		requestNode.put("left", "Footer links");
		requestNode.put("center", "Footer Mitte");
		requestNode.put("right", "Footer rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setFooters").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("sheet index (0) is out of range (no sheets)", responseNode.get(Executor.ERRORS).get(0).asText());

		Xls.activeWorkbook.createSheet();

		requestNode = mapper.createObjectNode();
		requestNode.put("left", "Footer links");
		requestNode.put("center", "Footer Mitte");
		requestNode.put("right", "Footer rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setFooters").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.SHEET.key(), SHEET0);
		requestNode.put("left", "Footer links");
		requestNode.put("center", "Footer Mitte");
		requestNode.put("right", "Footer rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setFooters").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.INDEX.key(), 0);
		requestNode.put("left", "Footer links");
		requestNode.put("center", "Footer Mitte");
		requestNode.put("right", "Footer rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setFooters").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.SHEET.key(), "Arbeitsblatt 1");
		requestNode.put("left", "Footer links");
		requestNode.put("center", "Footer Mitte");
		requestNode.put("right", "Footer rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setFooters").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("sheet with name 'Arbeitsblatt 1' does not exist", responseNode.get(Executor.ERRORS).get(0).asText());
		
		requestNode = mapper.createObjectNode();
		requestNode.put(Key.INDEX.key(), 1);
		requestNode.put("left", "Footer links");
		requestNode.put("center", "Footer Mitte");
		requestNode.put("right", "Footer rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setFooters").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("sheet index (1) is out of range (0..0)", responseNode.get(Executor.ERRORS).get(0).asText());

		requestNode = mapper.createObjectNode();
		requestNode.put(Key.INDEX.key(), 0);
		requestNode.put("left", "Footer links");
		requestNode.put("center", "Footer Mitte");
		requestNode.put("right", "Footer rechts");

		response = client.POST("http://localhost:4567/fsl/Xls.setFooters").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testJSONFormatting() throws JsonMappingException, JsonProcessingException
	{
		String value = "{\"amount\":287.30,\"currency\":\"CHF\",\"iban\":\"CH4431999123000889012\",\"reference\":\"000000000000000000000000000\",\"message\":\"Rechnungsnr. 10978 / Auftragsnr. 3987\",\"creditor\":{\"name\":\"Schreinerei Habegger & Söhne\",\"address_line_1\":\"Uetlibergstrasse 138\",\"address_line_2\":\"8045 Zürich\",\"country\":\"CH\"},\"debtor\":{\"name\":\"Simon Glarner\",\"address_line_1\":\"Bächliwis 55\",\"address_line_2\":\"8184 Bachenbülach\",\"country\":\"CH\"},\"format\":{\"graphics_format\":\"PDF\",\"output_size\":\"A4_PORTRAIT_SHEET\",\"language\":\"DE\"}}";
		JsonNode json = mapper.readTree(value);
		System.out.println(json.toPrettyString());
	}
}

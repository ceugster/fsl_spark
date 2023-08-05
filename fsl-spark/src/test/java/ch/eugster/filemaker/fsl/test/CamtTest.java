package ch.eugster.filemaker.fsl.test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeoutException;

import org.eclipse.jetty.client.api.ContentResponse;
import org.eclipse.jetty.client.util.StringContentProvider;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import ch.eugster.filemaker.fsl.Executor;
import ch.eugster.filemaker.fsl.camt.Camt;

public class CamtTest extends AbstractTest
{
	private String xmlFilename = "src/test/resources/xml/camt.054_P_CH0809000000450010065_1111204750_0_2022121623562233.xml";

	private String jsonFilename = "src/test/resources/json/camt.054_P_CH0809000000450010065_1111204750_0_2022121623562233.json";

	private String xmlContent;

	private String jsonContent;

	private String expectedXmlContent;
	
	private String expectedJsonContent;
	
	private static final String IDENTIFIER_KEY = "identifier";
	
	@BeforeEach
	public void before() throws IOException
	{
		File file = new File(xmlFilename);
		InputStream is = new FileInputStream(file);
		byte[] bytes = is.readAllBytes();
		expectedXmlContent = new String(bytes).replaceAll("\\s", "");
		is.close();
		file = new File(jsonFilename);
		is = new FileInputStream(file);
		bytes = is.readAllBytes();
		expectedJsonContent = new String(bytes).replaceAll("\\s", "");
		is.close();
		file = new File(xmlFilename);
		is = new FileInputStream(file);
		bytes = is.readAllBytes();
		xmlContent = new String(bytes);
		is.close();
		file = new File(jsonFilename);
		is = new FileInputStream(file);
		bytes = is.readAllBytes();
		jsonContent = new String(bytes);
		is.close();
	}
	
	@Test
	public void testConvertIllegalCamtXmlFile() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Camt.Key.XML_FILE.key(), "gigi");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Camt.convertCamt").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("'gigi' is not a valid xml file", responseNode.get(Executor.ERRORS).get(0).asText());
	}

	@Test
	public void testConvertIllegalCamtJsonFile() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Camt.Key.JSON_FILE.key(), "gigi");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Camt.convertCamt").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("'gigi' is not a valid json file", responseNode.get(Executor.ERRORS).get(0).asText());
	}

	@Test
	public void testConvertIllegalCamtXmlContent() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Camt.Key.XML_CONTENT.key(), "gigi");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Camt.convertCamt").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("xml content is not valid", responseNode.get(Executor.ERRORS).get(0).asText());
	}

	@Test
	public void testConvertIllegalCamtJsonContent() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Camt.Key.JSON_CONTENT.key(), "gigi");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Camt.convertCamt").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("json content is not valid", responseNode.get(Executor.ERRORS).get(0).asText());
	}

	@Test
	public void testConvertCamtXmlFile() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Camt.Key.XML_FILE.key(), xmlFilename);

		ContentResponse response = client.POST("http://localhost:4567/fsl/Camt.convertCamt").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(expectedJsonContent, responseNode.get(Executor.RESULT).asText().replaceAll("\\s", ""));
		assertEquals("camt.054.001.04", responseNode.get(IDENTIFIER_KEY).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testConvertCamtJsonFile() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Camt.Key.JSON_FILE.key(), jsonFilename);

		ContentResponse response = client.POST("http://localhost:4567/fsl/Camt.convertCamt").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(expectedXmlContent, responseNode.get(Executor.RESULT).asText().replaceAll("\\s", ""));
		assertEquals("camt.054.001.04", responseNode.get(IDENTIFIER_KEY).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testConvertCamtXmlContent() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Camt.Key.XML_CONTENT.key(), xmlContent);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Camt.convertCamt").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(expectedJsonContent, responseNode.get(Executor.RESULT).asText().replaceAll("\\s", ""));
		assertEquals("camt.054.001.04", responseNode.get(IDENTIFIER_KEY).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}

	@Test
	public void testConvertCamtJsonContent() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put(Camt.Key.JSON_CONTENT.key(), jsonContent);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Camt.convertCamt").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		assertEquals(expectedXmlContent, responseNode.get(Executor.RESULT).asText().replaceAll("\\s", ""));
		assertEquals("camt.054.001.04", responseNode.get(IDENTIFIER_KEY).asText());
		assertNull(responseNode.get(Executor.ERRORS));
	}
}

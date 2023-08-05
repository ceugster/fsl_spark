package ch.eugster.filemaker.fsl.test;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.sql.Connection;
import java.sql.Driver;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Objects;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeoutException;

import javax.sql.rowset.serial.SerialBlob;

import org.apache.commons.io.FileUtils;
import org.eclipse.jetty.client.api.ContentResponse;
import org.eclipse.jetty.client.util.StringContentProvider;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.BinaryNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import ch.eugster.filemaker.fsl.Executor;

public final class SqlTest extends AbstractTest
{
	private static Connection connection;
	
	private byte[] felix_gif;
	
	@BeforeAll
	public static void beforeAll() throws Exception
	{
		Class.forName("com.filemaker.jdbc.Driver").getConstructor().newInstance();
		AbstractTest.beforeAll();
	}
	
	@AfterAll
	protected static void afterAll() throws Exception
	{
		AbstractTest.afterAll();
	}
	
	@BeforeEach
	public void beforeEach() throws Exception
	{
		felix_gif = FileUtils.readFileToByteArray(new File("src/test/resources/gif/felix.gif"));
	}
	
	@AfterEach
	protected void afterEach() throws Exception
	{
		if (Objects.nonNull(connection) && !connection.isClosed())
		{
			connection.close();
		}
	}
	
	@Test
	public void getVersion() throws ClassNotFoundException, InstantiationException, IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException
	{
		Class<?> clazz = Class.forName("com.filemaker.jdbc.Driver");
		Driver driver = Driver.class.cast(clazz.getConstructor().newInstance());
		System.out.println(driver.getMajorVersion());
		System.out.println(driver.getMinorVersion());
		assertEquals("com.filemaker.jdbc.Driver", driver.getClass().getName());
	}
	
	@Test
	public void testConnectDisconnect() throws SQLException, InterruptedException, TimeoutException, ExecutionException, JsonMappingException, JsonProcessingException
	{
		String url = "jdbc:filemaker://localhost/Test";
		ObjectNode requestNode = mapper.createObjectNode();
		ObjectNode connectionNode = requestNode.objectNode();
		connectionNode.put("url", url);
		connectionNode.put("username", "christian");
		connectionNode.put("password", "ce_eu97");
		requestNode.set("connection", connectionNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Sql.connect").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		JsonNode statusNode = responseNode.get(Executor.STATUS);
		assertEquals("OK", statusNode.asText());
		JsonNode resultNode = responseNode.get("connection");
		assertEquals("jdbc:filemaker://localhost/Test?christian&ce_eu97", resultNode.asText());

		url = "jdbc:filemaker://localhost/Test";
		requestNode = mapper.createObjectNode();
		connectionNode = requestNode.objectNode();
		connectionNode.put("url", url);
		connectionNode.put("username", "christian");
		connectionNode.put("password", "ce_eu97");
		requestNode.set("connection", connectionNode);
		
		response = client.POST("http://localhost:4567/fsl/Sql.disconnect").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		responseNode = mapper.readTree(response.getContentAsString());
		statusNode = responseNode.get(Executor.STATUS);
		assertEquals("OK", statusNode.asText());
		resultNode = responseNode.get("connection");
		assertEquals("jdbc:filemaker://localhost/Test?christian&ce_eu97", resultNode.asText());
	}
	
	@Test
	public void testConnectFailWithIllegalUrl() throws SQLException, InterruptedException, TimeoutException, ExecutionException, JsonMappingException, JsonProcessingException
	{
		String url = "jdbc:filemaker://localhost/Kuckuck";
		ObjectNode requestNode = mapper.createObjectNode();
		ObjectNode connectionNode = requestNode.objectNode();
		connectionNode.put("url", url);
		connectionNode.put("username", "christian");
		connectionNode.put("password", "ce_eu97");
		requestNode.set("connection", connectionNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Sql.connect").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		JsonNode statusNode = responseNode.get(Executor.STATUS);
		assertEquals("Fehler", statusNode.asText());
		JsonNode errorsNode = responseNode.get(Executor.ERRORS);
		assertEquals("could not establish connection",  errorsNode.get(0).asText());
	}
	
	@Test
	public void testConnectFailWithoutUrl() throws SQLException, InterruptedException, TimeoutException, ExecutionException, JsonMappingException, JsonProcessingException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		ObjectNode connectionNode = requestNode.objectNode();
		connectionNode.put("username", "christian");
		connectionNode.put("password", "ce_eu97");
		requestNode.set("connection", connectionNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Sql.connect").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		JsonNode statusNode = responseNode.get(Executor.STATUS);
		assertEquals("Fehler", statusNode.asText());
		JsonNode errorsNode = responseNode.get(Executor.ERRORS);
		assertEquals("missing argument 'url'",  errorsNode.get(0).asText());
	}
	
	@Test
	public void testConnectWithoutPassword() throws SQLException, InterruptedException, TimeoutException, ExecutionException, JsonMappingException, JsonProcessingException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		ObjectNode connectionNode = requestNode.objectNode();
		connectionNode.put("url", "jdbc:filemaker://localhost/Test");
		connectionNode.put("username", "test");
		connectionNode.put("password", "");
		requestNode.set("connection", connectionNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Sql.connect").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		JsonNode statusNode = responseNode.get(Executor.STATUS);
		assertEquals("OK", statusNode.asText());
		JsonNode resultNode = responseNode.get("connection");
		assertEquals("jdbc:filemaker://localhost/Test?test&", resultNode.asText());
	}
	
	@Test
	public void testSelectWithWhereTypeString() throws InterruptedException, TimeoutException, ExecutionException, JsonMappingException, JsonProcessingException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		
		ObjectNode connectionNode = requestNode.objectNode();
		connectionNode.put("url", "jdbc:filemaker://localhost/Test");
		connectionNode.put("username", "test");
		connectionNode.put("password", "");
		requestNode.set("connection", connectionNode);
		
		requestNode.put("statement", "SELECT a.Id, a.Erstellt, a.ErstelltVon, a.Zahl, a.Datum, a.Zeit, a.Zeitstempel, GetAs(a.Container, DEFAULT), a.TextFormel, a.DatumFormel, a.IntegerFormel, a.DoubleFormel, a.ZeitFormel, a.ZeitstempelFormel  FROM Abfragen a WHERE a.TextFormel LIKE ?");
		
		ArrayNode parameterNode = requestNode.arrayNode();
		parameterNode.add("TextFormel");
		requestNode.set("parameter", parameterNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Sql.select").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		JsonNode statusNode = responseNode.get(Executor.STATUS);
		assertEquals("OK", statusNode.asText());
		JsonNode responseConnectionNode = responseNode.get("connection");
		assertEquals("jdbc:filemaker://localhost/Test?test&", responseConnectionNode.asText());
		JsonNode resultNode = responseNode.get("result");
		assertTrue(resultNode.isArray());
		assertEquals(1, resultNode.size());
		JsonNode rowNode = resultNode.get(0);
		assertEquals("06D88BB8-281E-44FB-A9DE-7AAD36691F0F", rowNode.get("Id").asText());
		assertEquals("01.07.2023 12:35:07", rowNode.get("Erstellt").asText());
		assertEquals("christian", rowNode.get("ErstelltVon").asText());
		assertEquals(1000.0, rowNode.get("Zahl").asDouble());
		assertEquals("21.10.1954", rowNode.get("Datum").asText());
		assertEquals("12:30:15", rowNode.get("Zeit").asText());
		byte[] bytes = null;
		try 
		{
			bytes = rowNode.get("Container").binaryValue();
		} 
		catch (IOException e) 
		{
		}
		assertArrayEquals(felix_gif, bytes);
		assertEquals("TextFormel", rowNode.get("TextFormel").asText());
		assertEquals("21.10.1954 10:30:27", rowNode.get("ZeitstempelFormel").asText());
		assertEquals("21.10.1954", rowNode.get("DatumFormel").asText());
		assertEquals("12:30:15", rowNode.get("ZeitFormel").asText());
		assertEquals("3.1427", rowNode.get("DoubleFormel").asText());
		assertEquals(2000.0, rowNode.get("IntegerFormel").asDouble());
	}

	@Test
	public void testSelectWithDifferentWhereTypes() throws InterruptedException, TimeoutException, ExecutionException, JsonMappingException, JsonProcessingException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		
		ObjectNode connectionNode = requestNode.objectNode();
		connectionNode.put("url", "jdbc:filemaker://localhost/Test");
		connectionNode.put("username", "test");
		connectionNode.put("password", "");
		requestNode.set("connection", connectionNode);
		
		requestNode.put("statement", "SELECT a.Id, a.Erstellt, a.ErstelltVon, a.Zahl, a.Datum, a.Zeit, a.Zeitstempel, GetAs(a.Container, DEFAULT), a.TextFormel, a.DatumFormel, a.IntegerFormel, a.DoubleFormel, a.ZeitFormel, a.ZeitstempelFormel  FROM Abfragen a WHERE a.Datum = ? AND a.Zeitstempel = ? AND a.Zeit = ? AND IntegerFormel = ? AND a.TextFormel = ?");
		
		ArrayNode parameterNode = requestNode.arrayNode();
		parameterNode.add("21.10.1954");
		parameterNode.add("21.10.1954 10:30:27");
		parameterNode.add("12:30:15");
		parameterNode.add(2000);
		parameterNode.add("TextFormel");
		requestNode.set("parameter", parameterNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Sql.select").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		JsonNode statusNode = responseNode.get(Executor.STATUS);
		assertEquals("OK", statusNode.asText());
		JsonNode responseConnectionNode = responseNode.get("connection");
		assertEquals("jdbc:filemaker://localhost/Test?test&", responseConnectionNode.asText());
		JsonNode resultNode = responseNode.get("result");
		assertTrue(resultNode.isArray());
		assertEquals(1, resultNode.size());
		JsonNode rowNode = resultNode.get(0);
		assertEquals("06D88BB8-281E-44FB-A9DE-7AAD36691F0F", rowNode.get("Id").asText());
		assertEquals("01.07.2023 12:35:07", rowNode.get("Erstellt").asText());
		assertEquals("christian", rowNode.get("ErstelltVon").asText());
		assertEquals(1000.0, rowNode.get("Zahl").asDouble());
		assertEquals("21.10.1954", rowNode.get("Datum").asText());
		assertEquals("12:30:15", rowNode.get("Zeit").asText());
		byte[] bytes = null;
		try 
		{
			bytes = rowNode.get("Container").binaryValue();
		} 
		catch (IOException e) 
		{
		}
		assertArrayEquals(felix_gif, bytes);
		assertEquals("TextFormel", rowNode.get("TextFormel").asText());
		assertEquals("21.10.1954 10:30:27", rowNode.get("ZeitstempelFormel").asText());
		assertEquals("21.10.1954", rowNode.get("DatumFormel").asText());
		assertEquals("12:30:15", rowNode.get("ZeitFormel").asText());
		assertEquals("3.1427", rowNode.get("DoubleFormel").asText());
		assertEquals(2000.0, rowNode.get("IntegerFormel").asDouble());
	}

	@Test
	public void testSelectWithWhereTypeZeitstempel() throws InterruptedException, TimeoutException, ExecutionException, JsonMappingException, JsonProcessingException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		
		ObjectNode connectionNode = requestNode.objectNode();
		connectionNode.put("url", "jdbc:filemaker://localhost/Test");
		connectionNode.put("username", "test");
		connectionNode.put("password", "");
		requestNode.set("connection", connectionNode);
		
		requestNode.put("statement", "SELECT a.Id FROM Abfragen a WHERE a.Zeitstempel = ?");
		
		ArrayNode parameterNode = requestNode.arrayNode();
		parameterNode.add("21.10.1954 10:30:27");
		requestNode.set("parameter", parameterNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Sql.select").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		JsonNode statusNode = responseNode.get(Executor.STATUS);
		assertEquals("OK", statusNode.asText());
		JsonNode responseConnectionNode = responseNode.get("connection");
		assertEquals("jdbc:filemaker://localhost/Test?test&", responseConnectionNode.asText());
		JsonNode resultNode = responseNode.get("result");
		assertTrue(resultNode.isArray());
		assertEquals(1, resultNode.size());
		JsonNode rowNode = resultNode.get(0);
		assertEquals("06D88BB8-281E-44FB-A9DE-7AAD36691F0F", rowNode.get("Id").asText());
	}

	@Test
	public void testSelectWithWhereTypeDate() throws InterruptedException, TimeoutException, ExecutionException, JsonMappingException, JsonProcessingException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		
		ObjectNode connectionNode = requestNode.objectNode();
		connectionNode.put("url", "jdbc:filemaker://localhost/Test");
		connectionNode.put("username", "test");
		connectionNode.put("password", "");
		requestNode.set("connection", connectionNode);
		
		requestNode.put("statement", "SELECT a.Id FROM Abfragen a WHERE a.Datum = ?");
		
		ArrayNode parameterNode = requestNode.arrayNode();
		parameterNode.add("21.10.1954");
		requestNode.set("parameter", parameterNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Sql.select").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		JsonNode statusNode = responseNode.get(Executor.STATUS);
		assertEquals("OK", statusNode.asText());
		JsonNode responseConnectionNode = responseNode.get("connection");
		assertEquals("jdbc:filemaker://localhost/Test?test&", responseConnectionNode.asText());
		JsonNode resultNode = responseNode.get("result");
		assertTrue(resultNode.isArray());
		assertEquals(1, resultNode.size());
		JsonNode rowNode = resultNode.get(0);
		assertEquals("06D88BB8-281E-44FB-A9DE-7AAD36691F0F", rowNode.get("Id").asText());
	}

	@Test
	public void testSelectWithWhereTypeTime() throws InterruptedException, TimeoutException, ExecutionException, JsonMappingException, JsonProcessingException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		
		ObjectNode connectionNode = requestNode.objectNode();
		connectionNode.put("url", "jdbc:filemaker://localhost/Test");
		connectionNode.put("username", "test");
		connectionNode.put("password", "");
		requestNode.set("connection", connectionNode);
		
		requestNode.put("statement", "SELECT a.Id FROM Abfragen a WHERE a.Zeit = ?");
		
		ArrayNode parameterNode = requestNode.arrayNode();
		parameterNode.add("12:30:15");
		requestNode.set("parameter", parameterNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Sql.select").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		JsonNode statusNode = responseNode.get(Executor.STATUS);
		assertEquals("OK", statusNode.asText());
		JsonNode responseConnectionNode = responseNode.get("connection");
		assertEquals("jdbc:filemaker://localhost/Test?test&", responseConnectionNode.asText());
		JsonNode resultNode = responseNode.get("result");
		assertTrue(resultNode.isArray());
		assertEquals(1, resultNode.size());
		JsonNode rowNode = resultNode.get(0);
		assertEquals("06D88BB8-281E-44FB-A9DE-7AAD36691F0F", rowNode.get("Id").asText());
	}
	
	@Test
	public void testInsertWithMissingPermission() throws InterruptedException, TimeoutException, ExecutionException, JsonMappingException, JsonProcessingException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		
		ObjectNode connectionNode = requestNode.objectNode();
		connectionNode.put("url", "jdbc:filemaker://localhost/Test");
		connectionNode.put("username", "test");
		connectionNode.put("password", "");
		requestNode.set("connection", connectionNode);
		
		requestNode.put("statement", "INSERT INTO Abfragen (Datum) VALUES (?)");
		
		ArrayNode parameterNode = requestNode.arrayNode();
		parameterNode.add("01.01.2022");
		requestNode.set("parameter", parameterNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Sql.insert").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		JsonNode statusNode = responseNode.get(Executor.STATUS);
		assertEquals(Executor.ERROR, statusNode.asText());
		JsonNode responseConnectionNode = responseNode.get("connection");
		assertEquals("jdbc:filemaker://localhost/Test?test&", responseConnectionNode.asText());
		assertNull(responseNode.get(Executor.RESULT));
	}

	@Test
	public void testInsertWithPermission() throws InterruptedException, TimeoutException, ExecutionException, IOException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		
		ObjectNode connectionNode = requestNode.objectNode();
		connectionNode.put("url", "jdbc:filemaker://localhost/Test");
		connectionNode.put("username", "christian");
		connectionNode.put("password", "ce_eu97");
		requestNode.set("connection", connectionNode);
		
		requestNode.put("statement", "INSERT INTO Abfragen (Container) VALUES (? AS 'file.gif')");
		
		File file = new File("src/test/resources/gif/felix.gif");
		byte[] bytes = FileUtils.readFileToByteArray(file);

		ArrayNode parameterNode = requestNode.arrayNode();
		BinaryNode binaryNode = parameterNode.binaryNode(bytes);
		parameterNode.add(binaryNode);
		parameterNode.add("06D88BB8-281E-44FB-A9DE-7AAD36691F0F");
		requestNode.set("parameter", parameterNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Sql.insert").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		JsonNode statusNode = responseNode.get(Executor.STATUS);
		assertEquals("OK", statusNode.asText());
		JsonNode responseConnectionNode = responseNode.get("connection");
		assertEquals("jdbc:filemaker://localhost/Test?christian&ce_eu97", responseConnectionNode.asText());
		JsonNode resultNode = responseNode.get("result");
		assertEquals(1, resultNode.asInt());
	}

	@Test
	public void testUpdateWithMissingPermission() throws InterruptedException, TimeoutException, ExecutionException, JsonMappingException, JsonProcessingException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		
		ObjectNode connectionNode = requestNode.objectNode();
		connectionNode.put("url", "jdbc:filemaker://localhost/Test");
		connectionNode.put("username", "test");
		connectionNode.put("password", "");
		requestNode.set("connection", connectionNode);
		
		requestNode.put("statement", "UPDATE Abfragen SET Datum = ?");
		
		ArrayNode parameterNode = requestNode.arrayNode();
		parameterNode.add("01.01.2022");
		requestNode.set("parameter", parameterNode);
		
		ContentResponse response = client.POST("http://localhost:4567/fsl/Sql.update").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		JsonNode statusNode = responseNode.get(Executor.STATUS);
		assertEquals("Fehler", statusNode.asText());
		JsonNode responseConnectionNode = responseNode.get("connection");
		assertEquals("jdbc:filemaker://localhost/Test?test&", responseConnectionNode.asText());
		assertNull(responseNode.get("result"));
	}

	@Test
	public void testUpdateWithPermission() throws InterruptedException, TimeoutException, ExecutionException, IOException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		
		ObjectNode connectionNode = requestNode.objectNode();
		connectionNode.put("url", "jdbc:filemaker://localhost/Test");
		connectionNode.put("username", "christian");
		connectionNode.put("password", "ce_eu97");
		requestNode.set("connection", connectionNode);
		
		requestNode.put("statement", "UPDATE Abfragen SET Container = ? AS 'felix.gif' WHERE Id = ?");
		
		File file = new File("src/test/resources/gif/felix.gif");
		byte[] bytes = FileUtils.readFileToByteArray(file);

		ArrayNode parameterNode = requestNode.arrayNode();
		BinaryNode binaryNode = parameterNode.binaryNode(bytes);
		parameterNode.add(binaryNode);
		parameterNode.add("06D88BB8-281E-44FB-A9DE-7AAD36691F0F");
		requestNode.set("parameter", parameterNode);
		System.out.println(requestNode.toString());
		ContentResponse response = client.POST("http://localhost:4567/fsl/Sql.update").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		JsonNode statusNode = responseNode.get(Executor.STATUS);
		assertEquals("OK", statusNode.asText());
		JsonNode responseConnectionNode = responseNode.get("connection");
		assertEquals("jdbc:filemaker://localhost/Test?christian&ce_eu97", responseConnectionNode.asText());
		JsonNode resultNode = responseNode.get("result");
		assertEquals(1, resultNode.asInt());
	}

	@Test
	public void test() throws InterruptedException, TimeoutException, ExecutionException, IOException, ClassNotFoundException, SQLException
	{
		File file = new File("src/test/resources/gif/felix.gif");
		byte[] bytes = FileUtils.readFileToByteArray(file);
		Class.forName("com.filemaker.jdbc.Driver");
		Connection connection = DriverManager.getConnection("jdbc:filemaker://localhost/Test", "christian", "ce_eu97");
		PreparedStatement statement = connection.prepareStatement("UPDATE Abfragen SET Container = PutAs(?, 'GIFf') WHERE Id = ?");
		statement.setBytes(1, bytes);
		statement.setString(2, "06D88BB8-281E-44FB-A9DE-7AAD36691F0F");
		try
		{
			int result = statement.executeUpdate();
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}

	@Test
	public void testParameter()
	{
		// gültige/ungültige Parametertypen
		// gültige/ungültige Anzahl Parameter
		// container als parameter
	}
}

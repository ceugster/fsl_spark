package ch.eugster.filemaker.fsl.test;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.Objects;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeoutException;

import org.eclipse.jetty.client.api.ContentResponse;
import org.eclipse.jetty.client.util.StringContentProvider;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.JsonNode;

import ch.eugster.filemaker.fsl.Executor;
import ch.eugster.filemaker.fsl.Fsl;

public class FslTest extends AbstractTest
{
	private Fsl fsl;
	
	@BeforeEach
	public void beforeEach() throws Exception
	{
		if (Objects.isNull(fsl))
		{
			fsl = new Fsl();
		}
	}

	@AfterEach
	public void afterEach() throws Exception
	{
		fsl = null;
	}

	@Test
	public void testFslWithoutModule() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ContentResponse response = client.POST("http://localhost:4567/fsl").
				header("Content-Type", "application/json").
				content(new StringContentProvider("{}")).
				accept("application/json").
				send();
		
		assertEquals("<html><body><h2>404 Not found</h2></body></html>", response.getContentAsString());
	}

	@Test
	public void testFslWithoutCommand() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ContentResponse response = client.POST("http://localhost:4567/fsl").
				header("Content-Type", "application/json").
				content(new StringContentProvider("{}")).
				accept("application/json").
				send();
		
		assertEquals("<html><body><h2>404 Not found</h2></body></html>", response.getContentAsString());
	}

	@Test
	public void testFslWithWrongModule() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ContentResponse response = client.POST("http://localhost:4567/fsl/Schmock").
				header("Content-Type", "application/json").
				content(new StringContentProvider("{}")).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("missing command", responseNode.get(Executor.ERRORS).get(0).asText());
	}

	@Test
	public void testFslWithWrongCommand() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ContentResponse response = client.POST("http://localhost:4567/fsl/Xls.TschaTscha").
				header("Content-Type", "application/json").
				content(new StringContentProvider("{}")).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("Invalid command 'TschaTscha'", responseNode.get(Executor.ERRORS).get(0).asText());
	}

	@Test
	public void testFslWithWrongModuleAndCommand() throws JsonMappingException, JsonProcessingException, InterruptedException, TimeoutException, ExecutionException
	{
		ContentResponse response = client.POST("http://localhost:4567/fsl/Schmock.TschaTscha").
				header("Content-Type", "application/json").
				content(new StringContentProvider("{}")).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("illegal module", responseNode.get(Executor.ERRORS).get(0).asText());
	}
}

package ch.eugster.filemaker.fsl.test;

import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

import org.eclipse.jetty.client.HttpClient;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;

import com.fasterxml.jackson.databind.ObjectMapper;

import ch.eugster.filemaker.fsl.Executor;
import ch.eugster.filemaker.fsl.Fsl;

public abstract class AbstractTest
{
	protected static ObjectMapper mapper = new ObjectMapper();

	protected static HttpClient client;

	protected static Fsl fsl;
	
	protected static Map<String, Executor> executors = new HashMap<String, Executor>();
	
	@BeforeAll
	protected static void beforeAll() throws Exception
	{
		Fsl fsl = Fsl.fsls.get(System.getProperty("user.name"));
		if (Objects.isNull(fsl))
		{
			fsl = new Fsl();
			Fsl.fsls.put(System.getProperty("user.name"), fsl);
		}
		if (Objects.isNull(client))
		{
			client = new HttpClient();
		}
		client.start();
	}
	
	@AfterAll
	protected static void afterAll() throws Exception
	{
		client.stop();
	}
	
}

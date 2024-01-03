package ch.eugster.filemaker.fsl;

import static spark.Spark.post;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Objects;
import java.util.Properties;
import java.util.logging.LogManager;

import org.reflections.Reflections;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.slf4j.event.Level;

import com.fasterxml.jackson.core.JsonParseException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import spark.servlet.SparkApplication;

public class Fsl implements SparkApplication
{
	public static String VERSION = "1.1.1";
	
	public static Map<String, Fsl> fsls = new HashMap<String, Fsl>();
	
	private Map<String, Executor> executors = new HashMap<String, Executor>();

	private static ObjectMapper mapper = new ObjectMapper();

	private static Reflections reflections = new Reflections("ch.eugster.filemaker.fsl");

	private static Logger logger = LoggerFactory.getLogger(Fsl.class);
	
	private static Path fslPath = Paths.get(System.getProperty("user.home"), ".fsl");
	
	public static void main(String[] args)
	{
//		readConfiguration();
		initializeLogging();
		Fsl fsl = fsls.get(System.getProperty("user.name"));
		if (Objects.isNull(fsl))
		{
			logger.info("add new fsl for '" + System.getProperty("user.name") + "'");
			fsl = new Fsl();
			fsls.put(System.getProperty("user.name"), fsl);
		}
	}
	
	public static String execute(String command, String parameters)
	{
		logger.info("execute command '" + command + "' with parameters '" + parameters + "'");
		Fsl fsl = getFsl(System.getProperty("user.name"));
		return fsl.doExecute(command, parameters);
	}
	
	public static Fsl getFsl(String key)
	{
		Fsl fsl = fsls.get(System.getProperty("user.name"));
		if (Objects.isNull(fsl))
		{
			fsl = new Fsl();
			fsls.put(System.getProperty("user.name"), fsl);
		}
		return fsl;
	}
	
	public static void log(Level level, String message)
	{
		logger.info(message);
	}
	
	public static void releaseFsls()
	{
		fsls.clear();
	}
	
	public Fsl()
	{
		init();
	}
	
	@Override
	public void init()
	{
		post("/fsl/:command", (request, response) -> 
		{
			ObjectNode responseNode = mapper.createObjectNode();
			String command = request.params(":command");
			if (Objects.nonNull(command))
			{
				if (command.equals("Fsl.version"))
				{
					response.status(200);
					responseNode.put(Executor.STATUS, Executor.OK);
					responseNode.put("version", VERSION);
				}
				else
				{
					String[] commandParts = command.split("[.]");
					if (commandParts.length == 0)
					{
						response.status(503);
						ArrayNode errors = responseNode.arrayNode();
						errors.add("missing module and command");
						responseNode.set(Executor.ERRORS, errors);
					}
					else if (commandParts.length == 1)
					{
						response.status(503);
						ArrayNode errors = responseNode.arrayNode();
						errors.add("missing command");
						responseNode.set(Executor.ERRORS, errors);
					}
					else if (commandParts.length == 2)
					{
						Executor executor = getExecutor(commandParts[0]);
						if (Objects.nonNull(executor))
						{
							try 
							{
								JsonNode jsonNode = mapper.readTree(request.body());
								if (jsonNode.isObject())
								{
									ObjectNode requestNode = ObjectNode.class.cast(jsonNode);
									responseNode = executor.execute(commandParts[1], requestNode, responseNode);
									if (responseNode.has(Executor.ERRORS))
									{
										response.status(503);
									}
									else
									{
										response.status(200);
										response.type("application/json");
									}
								}
								else
								{
									response.status(503);
									ArrayNode errors = responseNode.arrayNode();
									errors.add("missing or illegal argument");
									responseNode.set(Executor.ERRORS, errors);
								}
							}
							catch (JsonParseException jpe) 
							{
								response.status(503);
								ArrayNode errors = responseNode.arrayNode();
								errors.add("illegal arguments");
								responseNode.set(Executor.ERRORS, errors);
							}
						} 
						else
						{
							response.status(503);
							ArrayNode errors = responseNode.arrayNode();
							errors.add("illegal module");
							responseNode.set(Executor.ERRORS, errors);
						}
					}
					else
					{
						response.status(503);
						ArrayNode errors = responseNode.arrayNode();
						errors.add("illegal module and command");
						responseNode.set(Executor.ERRORS, errors);
					}
				}
			}
			else
			{
				response.status(503);
				ArrayNode errors = responseNode.arrayNode();
				errors.add("missing module and command");
				responseNode.set(Executor.ERRORS, errors);
			}
			return responseNode.put(Executor.STATUS, Objects.isNull(responseNode.get("errors")) && response.status() == 200 ? Executor.OK : Executor.ERROR).toString();
		});
	}
	
	private String doExecute(String command, String parameters)
	{
		logger.info("Build argument node from '" + parameters + "'");
		ObjectNode requestNode = null;
		ObjectNode responseNode = mapper.createObjectNode();
		try
		{
			requestNode = ObjectNode.class.cast(mapper.readTree(parameters));
			if (Objects.nonNull(command) && !command.trim().isEmpty())
			{
				String[] commandParts = command.split("[.]");
				if (commandParts.length == 2)
				{
					logger.info("Get or create executor '{}'", commandParts[0]);
					Executor executor = getExecutor(commandParts[0].trim());
					if (Objects.nonNull(executor))
					{
						logger.info("Execute '" + commandParts[1] + "'");
						executor.execute(commandParts[1].trim(), requestNode, responseNode);
					}
					else
					{
						logger.error("illegal module '" + commandParts[0] + "'");
						addErrorMessage(responseNode, "invalid module '" + commandParts[0] + "'");
					}
				}
				else
				{
					logger.error("illegal command '{}'", command);
					addErrorMessage(responseNode, "invalid command '" + command + "'");
				}
			}
			else
			{
				logger.error("Missing command");
				addErrorMessage(responseNode, "missing command");
			}
		}
		catch (Exception e)
		{
			logger.error("An error occurred while building the argument node ({})", e.getLocalizedMessage());
			addErrorMessage(responseNode, "illegal argument '" + parameters + "'");
		}

		return responseNode.put(Executor.STATUS, Objects.isNull(responseNode.get(Executor.ERRORS)) ? Executor.OK : Executor.ERROR).toString();
	}
	
	public boolean addErrorMessage(ObjectNode responseNode, String message)
	{
		ArrayNode errors = ArrayNode.class.cast(responseNode.get(Executor.ERRORS));
		if (Objects.isNull(errors))
		{
			errors = responseNode.arrayNode();
			responseNode.set(Executor.ERRORS, errors);
		}
		if (errors.findValuesAsText(message).size() == 0)
		{
			errors.add(message);
		}
		return false;
	}
	
	public Executor getExecutor(String executorName)
	{
		Executor executor = null;
		executor = executors.get(executorName);
		if (Objects.isNull(executor))
		{
			try
			{
				Iterator<Class<? extends Executor>> clazzes = reflections.getSubTypesOf(Executor.class).iterator();
				while (clazzes.hasNext())
				{
					Class<? extends Executor> clazz = clazzes.next();
					if (clazz.getSimpleName().equals(executorName))
					{
						executor = clazz.getConstructor().newInstance();
						executors.put(executorName, executor);
						break;
					}
				}
			}
			catch (Exception e)
			{
			}
		}
		return executor;
	}
	
//	private static void readConfiguration()
//	{
//		Path cfgPath = Paths.get(fslPath.toString(), "fsl.cfg");
//		File file = new File(cfgPath.toString());
//		Reader reader = null;
//		try 
//		{
//			reader = new FileReader(file);
//			properties = new Properties();
//			properties.setProperty("ch.eugster.filemaker.fsl.url", "http://localhost");
//			properties.setProperty("ch.eugster.filemaker.fsl.port", "4567");
//			properties.load(reader);
//		} 
//		catch (FileNotFoundException e) 
//		{
//			writeConfiguration();
//		} 
//		catch (IOException e) 
//		{
//		}
//		finally
//		{
//			try
//			{
//				reader.close();
//			}
//			catch (IOException e) 
//			{
//			}
//		}
//	}
//
//	private static void writeConfiguration()
//	{
//		Path cfgPath = Paths.get(fslPath.toString(), "fsl.cfg");
//		File file = new File(cfgPath.toString());
//		Writer writer = null;
//		try 
//		{
//			writer = new FileWriter(file);
//			properties.store(writer, "Do not change the content of this file manually");
//		} 
//		catch (IOException e) 
//		{
//		}
//		finally
//		{
//			try
//			{
//				writer.close();
//			}
//			catch (IOException e) 
//			{
//			}
//		}
//	}
	
	private static void initializeLogging()
	{
		Path cfgPath = Paths.get(fslPath.toString(), "fsl-log.cfg");
		Path logPath = Paths.get(fslPath.toString(), "fsl.log");
		System.setProperty("java.util.logging.config.file", cfgPath.toString());
		if (logPath.getParent().toFile().exists())
		{
			try
			{
				logPath.getParent().toFile().mkdirs();
			} 
			catch (Exception e)
			{
			}
		}
		if (!cfgPath.toFile().exists())
		{
			Properties properties = new Properties();
			properties.setProperty("handlers", "java.util.logging.FileHandler, java.util.logging.ConsoleHandler");
			properties.setProperty("java.util.logging.FileHandler.pattern", "%h/.fsl/fsl.log");
			properties.setProperty("java.util.logging.FileHandler.formatter", "java.util.logging.SimpleFormatter");
			properties.setProperty("java.util.logging.FileHandler.level", "INFO");
			properties.setProperty("java.util.logging.ConsoleHandler.level", "INFO");
			properties.setProperty("ch.eugster.filemaker.fsl.log", Boolean.toString(false));
			OutputStream os = null;
			try
			{
				os = new FileOutputStream(cfgPath.toFile());
				properties.store(os, "Please do not change the content of this file");
			}
			catch (Exception e)
			{
			}
			finally
			{
				if (Objects.nonNull(os)) 
				{
					try
					{
						os.close();
					}
					catch (Exception e)
					{
					}
				}
			}
		}
		try
		{
			InputStream is = new FileInputStream(cfgPath.toFile());
			LogManager.getLogManager().readConfiguration(is);
		}
		catch (Exception e)
		{
		}
	}
	
}

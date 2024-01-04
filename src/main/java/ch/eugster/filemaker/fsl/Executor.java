package ch.eugster.filemaker.fsl;

import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.Objects;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

public abstract class Executor
{
	public static final String STATUS = "status";
	
	public static final String ERRORS = "errors";
	
	public static final String OK = "OK";

	public static final String ERROR = "Fehler";

	public static final String RESULT = "result";

	protected static ObjectMapper mapper = new ObjectMapper();
	
	private ObjectNode requestNode;
	
	private ObjectNode responseNode;
	
	public ObjectNode execute(String command, ObjectNode requestNode, ObjectNode responseNode)
	{
		this.requestNode = requestNode;
		this.responseNode = responseNode;
		
		boolean found = false;
		Method[] methods = this.getClass().getDeclaredMethods();
		for (Method method : methods)
		{
			if (method.getModifiers() == Modifier.PUBLIC)
			{
				if (method.getName().equals(command))
				{
//					Class<?>[] types = method.getParameterTypes();
					found = true;
//					for (Class<?> type : types)
//					{
//						if (!String.class.equals(type))
//						{
//							found = false;
//						}
//					}
//					if (found)
//					{
						try
						{
							method.invoke(this);
							responseNode.put(Executor.STATUS, responseNode.has(Executor.ERRORS) ? Executor.ERROR : Executor.OK);
						}
						catch (Exception e)
						{
							ArrayNode errors = ArrayNode.class.cast(responseNode.get(Executor.ERRORS));
							if (Objects.isNull(errors))
							{
								errors = responseNode.arrayNode();
								responseNode.set(Executor.ERRORS, errors);
							}
							errors.add(Objects.isNull(e.getLocalizedMessage()) ? e.getClass().getName() : e.getLocalizedMessage());
						}
						break;
//					}
				}
			}
		}
		if (!found)
		{
			addErrorMessage("Invalid command '" + command + "'");
		}
		return responseNode;
	}
	
	protected boolean createRequestNode(String request)
	{
		boolean result = true;
		if (Objects.nonNull(request) && !request.trim().isEmpty())
		{
			try 
			{
				JsonNode _requestNode = mapper.readTree(request);
				requestNode = ObjectNode.class.cast(_requestNode);
				responseNode = mapper.createObjectNode();
			} 
			catch (JsonMappingException e) 
			{
				result = addErrorMessage("cannot map 'request': illegal json format");
			} 
			catch (JsonProcessingException e) 
			{
				result = addErrorMessage("cannot process 'request': illegal json format");
			}
			catch (ClassCastException e)
			{
				result = addErrorMessage("cannot cast 'request': illegal json format");
			}
		}
		else
		{
			result = addErrorMessage("missing argument 'request'");
		}
		return result;
	}

	protected ObjectNode getRequestNode()
	{
		return requestNode;
	}

	protected ObjectNode getResponseNode()
	{
		if (Objects.isNull(responseNode))
		{
			responseNode = mapper.createObjectNode();
		}
		return responseNode;
	}
	
	protected String getResponse()
	{
		String response = getResponseNode().put(Executor.STATUS, getResponseNode().has(Executor.ERRORS) ? Executor.ERROR : Executor.OK).toString();
		requestNode = null;
		responseNode = null;
		return response;
	}
	
	protected boolean addErrorMessage(String message)
	{
		ArrayNode errors = ArrayNode.class.cast(responseNode.get(Executor.ERRORS));
		if (Objects.isNull(errors))
		{
			errors = responseNode.arrayNode();
			responseNode.set(Executor.ERRORS, errors);
		}
		errors.add(message);
		return false;
	}
}

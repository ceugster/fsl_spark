package ch.eugster.filemaker.fsl;

import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.lang.reflect.Parameter;
import java.util.Objects;

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
	
	protected ObjectNode requestNode;
	
	protected ObjectNode responseNode;
	
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
				if (method.getName().equals(command) && method.getParameterCount() == 2)
				{
					Class<?>[] types = method.getParameterTypes();
					found = true;
					for (Class<?> type : types)
					{
						if (!ObjectNode.class.equals(type))
						{
							found = false;
						}
					}
					if (found)
					{
						try
						{
							method.invoke(this, requestNode, responseNode);
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
					}
				}
			}
		}
		if (!found)
		{
			addErrorMessage("Invalid command '" + command + "'");
		}
		return responseNode;
	}
	
	protected boolean addErrorMessage(String message)
	{
		if (Objects.isNull(responseNode))
		{
			responseNode = new ObjectMapper().createObjectNode();
			
		}
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

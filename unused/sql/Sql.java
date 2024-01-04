package ch.eugster.filemaker.fsl.sql;

import java.lang.reflect.InvocationTargetException;
import java.sql.Blob;
import java.sql.Connection;
import java.sql.Date;
import java.sql.Driver;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Time;
import java.sql.Timestamp;
import java.sql.Types;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

import org.apache.commons.codec.binary.Base64;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import ch.eugster.filemaker.fsl.Executor;

public class Sql extends Executor
{
	private static Logger logger = LoggerFactory.getLogger(Sql.class);
	
	private Map<String, Connection> connections = new HashMap<String, Connection>();
	
	private Connection connection;
	
	public boolean connect(ObjectNode requestNode, ObjectNode responseNode)
	{
		connection = getConnection(requestNode, responseNode);
		return Objects.nonNull(connection);
	}
	
	public boolean disconnect(ObjectNode requestNode, ObjectNode responseNode)
	{
		boolean result = true;
		ConnectionDetails details = getConnectionDetails(requestNode);
		Connection connection = connections.remove(details.key());
		if (Objects.nonNull(connection))
		{
			try
			{
				connection.close();
				if (connection == this.connection)
				{
					this.connection = null;
				}
				responseNode.put("connection", details.key());
			}
			catch (SQLException e)
			{
				result = addErrorMessage("could not close connection '" + details.key() + "'");
			}
		}
		else
		{
			result = addErrorMessage("connection does not exist '" + details.key() + "'");
		}
		return result;
	}
	
	public boolean select(ObjectNode requestNode, ObjectNode responseNode)
	{
		boolean result = true;
		ConnectionDetails details = getConnectionDetails(requestNode);
		if (Objects.nonNull(details))
		{
			if (connect(requestNode, responseNode))
			{
				try
				{
					PreparedStatement preparedStatement = createPreparedStatement(requestNode);
					ArrayNode resultNode = responseNode.arrayNode();
					ResultSet response = preparedStatement.executeQuery();
					ResultSetMetaData rsmd = response.getMetaData();
					while (response.next())
					{
						ObjectNode rowNode = resultNode.objectNode();
						int cols = rsmd.getColumnCount();
						for (int col = 1; col < cols + 1; col++)
						{
							String key = rsmd.getColumnName(col);
							switch (rsmd.getColumnType(col))
							{
								case Types.VARCHAR:
								{
									String value = response.getString(col);
									rowNode.put(key, value);
									break;
								}
								case Types.TIMESTAMP:
								{
									Timestamp ts = response.getTimestamp(col);
									Date date = new Date(ts.getTime());
									DateFormat df = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss");
									String value = df.format(date);
									rowNode.put(key, value);
									break;
								}
								case Types.DATE:
								{
									Date d = response.getDate(col);
									Date date = new Date(d.getTime());
									DateFormat df = new SimpleDateFormat("dd.MM.yyyy");
									String value = df.format(date);
									rowNode.put(key, value);
									break;
								}
								case Types.TIME:
								{
									Time d = response.getTime(col);
									Date date = new Date(d.getTime());
									DateFormat df = new SimpleDateFormat("HH:mm:ss");
									String value = df.format(date);
									rowNode.put(key, value);
									break;
								}
								case Types.INTEGER:
								{
									int value = response.getInt(col);
									rowNode.put(key, value);
									break;
								}
								case Types.DOUBLE:
								{
									double value = response.getDouble(col);
									rowNode.put(key, value);
									break;
								}
								case Types.BLOB:
								{
									Blob value = response.getBlob(col);
									rowNode.put(key, value.getBytes(col, Long.valueOf(value.length()).intValue()));
									break;
								}
								case Types.LONGVARBINARY:
								{
									byte[] bytes = response.getBytes(col);
									rowNode.put(key, bytes);
									break;
								}
								default:
								{
									throw new IllegalArgumentException(key);
								}
							}
						}
						resultNode.add(rowNode);
						rowNode = resultNode.objectNode();
					}
					responseNode.set("result", resultNode);
				}
				catch (SQLException e)
				{
					addErrorMessage("");
				}
			}
		}
		return result;
	}

	public boolean update(ObjectNode requestNode, ObjectNode responseNode)
	{
		boolean result = true;
		ConnectionDetails details = getConnectionDetails(requestNode);
		if (Objects.nonNull(details))
		{
			if (connect(requestNode, responseNode))
			{
				PreparedStatement preparedStatement = createPreparedStatement(requestNode);
				if (Objects.nonNull(preparedStatement))
				{
					try
					{
						int response = preparedStatement.executeUpdate();
						responseNode.put("result", response);
					}
					catch (SQLException e)
					{
						if (e.getSQLState().equals("HY000"))
						{
							result = addErrorMessage("no write permission");
						}
						else
						{
							result = addErrorMessage(e.getLocalizedMessage());
						}
					}
				}
			}
			else
			{
				result = addErrorMessage("could not connect to '" + details.key() + "'");
			}
		}
		return result;
	}

	public boolean insert(ObjectNode requestNode, ObjectNode responseNode)
	{
		boolean result = true;
		ConnectionDetails details = getConnectionDetails(requestNode);
		if (Objects.nonNull(details))
		{
			if (connect(requestNode, responseNode))
			{
				PreparedStatement preparedStatement = createPreparedStatement(requestNode);
				if (Objects.nonNull(preparedStatement))
				{
					try
					{
						int response = preparedStatement.executeUpdate();
						responseNode.put("result", response);
					}
					catch (SQLException e)
					{
						if (e.getSQLState().equals("HY000"))
						{
							result = addErrorMessage("no write permission");
						}
						else
						{
							result = addErrorMessage(e.getLocalizedMessage());
						}
					}
				}
			}
			else
			{
				result = addErrorMessage("could not connect to '" + details.key() + "'");
			}
		}
		return result;
	}

	private PreparedStatement createPreparedStatement(ObjectNode requestNode)
	{
		PreparedStatement preparedStatement = null;
		JsonNode statementNode = requestNode.findPath("statement");
		if (statementNode.isTextual())
		{
			try
			{
				String statement = statementNode.asText();
				int numberOfQuestionMarks = getNumberOfQuestionMarks(statement);
				preparedStatement = connection.prepareStatement(statement);
				JsonNode parameterNode = requestNode.findPath("parameter");
				if (parameterNode.isArray() && numberOfQuestionMarks > 0)
				{
					preparedStatement = fillParameters(preparedStatement, ArrayNode.class.cast(parameterNode), numberOfQuestionMarks);
					
				}
			}
			catch ( SQLException e)
			{
				addErrorMessage("illegal argument 'statement'");
			}
		}
		return preparedStatement;
	}
	
	private PreparedStatement fillParameters(PreparedStatement preparedStatement, ArrayNode parameterNode, int numberOfQuestionMarks)
	{
		if (parameterNode.size() < numberOfQuestionMarks)
		{
			addErrorMessage("illegal number of parameters");
		}
		else
		{
			JsonNode currentNode = null;
			try
			{
				for (int i = 0; i < numberOfQuestionMarks; i++)
				{
					currentNode = parameterNode.get(i);
					if (currentNode.isTextual())
					{
						try
						{
							Timestamp timestamp = getTimestamp(currentNode);
							preparedStatement.setTimestamp(i + 1, timestamp);
						}
						catch (ParseException e0)
						{
							try
							{
								Date date = getDate(currentNode);
								preparedStatement.setDate(i + 1, date);
							}
							catch (ParseException e1)
							{
								try
								{
									Time time = getTime(currentNode);
									preparedStatement.setTime(i + 1, time);
								}
								catch (ParseException e2)
								{
									preparedStatement.setString(i + 1, currentNode.asText());
								}
							}
						}
					}
					else if (currentNode.isInt())
					{
						preparedStatement.setInt(i + 1, currentNode.asInt());
					}
					else if (currentNode.isLong())
					{
						preparedStatement.setLong(i + 1, currentNode.asLong());
					}
					else if (currentNode.isDouble())
					{
						preparedStatement.setDouble(i + 1, currentNode.asDouble());
					}
					else
					{
						System.out.println();
					}
				}
			}
			catch (SQLException e)
			{
				addErrorMessage("illegal parameter '" + currentNode.asText() + "'");
			}
		}
		return preparedStatement;
	}
	
	private ConnectionDetails getConnectionDetails(ObjectNode requestNode)
	{
		ConnectionDetails details = null;
		JsonNode connectionNode = requestNode.findPath("connection");
		if (connectionNode.isMissingNode())
		{
			addErrorMessage("missing argument 'connection'");
		}
		else if (connectionNode.isTextual())
		{
			details = new ConnectionDetails(connectionNode.asText());
		}
		else if (connectionNode.isObject())
		{
			JsonNode urlNode = connectionNode.findPath("url");
			if (urlNode.isMissingNode())
			{
				addErrorMessage("missing argument 'url'");
			}
			else if (urlNode.isTextual())
			{
				String url = urlNode.asText();
				JsonNode usernameNode = connectionNode.findPath("username");
				if (usernameNode.isMissingNode())
				{
					addErrorMessage("missing argument 'username'");
				}
				else if (usernameNode.isTextual())
				{
					String username = usernameNode.asText();
					JsonNode passwordNode = connectionNode.findPath("password");
					if (passwordNode.isTextual())
					{
						String password = passwordNode.asText();
						details = new ConnectionDetails(url, username, password);
					}
				}
				else
				{
					addErrorMessage("illegal argument 'username'");
				}
			}
			else
			{
				addErrorMessage("illegal argument 'url'");
			}
		}
		else
		{
			addErrorMessage("illegal argument 'connection'");
		}
		return details;
	}
	
	private Connection getConnection(ObjectNode requestNode, ObjectNode responseNode)
	{
		ConnectionDetails details = getConnectionDetails(requestNode);
		if (Objects.isNull(connections.get(details.key())))
		{
			connection = createConnection(requestNode, responseNode);
		}
		else
		{
			responseNode.put("connection", details.key());
		}
		return connection;
	}
	
	private Connection createConnection(ObjectNode requestNode, ObjectNode responseNode)
	{
		ConnectionDetails details = getConnectionDetails(requestNode);
		try
		{
			Class<?> clazz = Class.forName("com.filemaker.jdbc.Driver");
			Driver.class.cast(clazz.getConstructor().newInstance());
			connection = DriverManager.getConnection(details.url(), details.username(), details.password());
			connections.put(details.key(), connection);
			responseNode.put("connection", details.key());
			
		}
		catch (ClassNotFoundException e)
		{
			addErrorMessage("missing filemaker jdbc driver");
		}
		catch (SQLException e)
		{
			addErrorMessage("could not establish connection");
		} 
		catch (InstantiationException e) 
		{
			addErrorMessage("illegal driver");
		} 
		catch (IllegalAccessException e) 
		{
			addErrorMessage("illegal driver");
		} 
		catch (IllegalArgumentException e) 
		{
			addErrorMessage("illegal driver");
		} 
		catch (InvocationTargetException e) 
		{
			addErrorMessage("illegal driver");
		} 
		catch (NoSuchMethodException e) 
		{
			addErrorMessage("illegal driver");
		} 
		catch (SecurityException e) 
		{
			addErrorMessage("illegal driver");
		}
		return connection;
	}
	
	private int getNumberOfQuestionMarks(String query)
	{
		return Long.valueOf(query.chars().filter(ch -> ch == '?').count()).intValue();
	}
	
	class ConnectionDetails
	{
		private String url;
		
		private String username;
		
		private String password;
		
		ConnectionDetails(String url, String username, String password)
		{
			this.url = url;
			this.username = username;
			this.password = password;
		}
		
		ConnectionDetails(String key)
		{
			this.url = key.substring(0, key.indexOf("?"));
			this.username = key.substring(key.indexOf("?"), key.indexOf("&"));
			this.password = key.substring(key.indexOf("&"), key.length());
			System.out.println(key);
		}
		
		public String key()
		{
			return url + "?" + username + "&" + password;
		}
		
		public String url()
		{
			return this.url;
		}

		public String username()
		{
			return this.username;
		}

		public String password()
		{
			return this.password;
		}
	}
	
	private java.sql.Timestamp getTimestamp(JsonNode parameterNode) throws ParseException
	{
		String value = parameterNode.asText();
		DateFormat df = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss");
		java.sql.Timestamp date = new java.sql.Timestamp(df.parse(value).getTime());
		return date;
	}

	private java.sql.Date getDate(JsonNode parameterNode) throws ParseException
	{
		String value = parameterNode.asText();
		DateFormat df = new SimpleDateFormat("dd.MM.yyyy");
		java.sql.Date date = new java.sql.Date(df.parse(value).getTime());
		return date;
	}

	private java.sql.Time getTime(JsonNode parameterNode) throws ParseException
	{
		String value = parameterNode.asText();
		DateFormat df = new SimpleDateFormat("HH:mm:ss");
		java.sql.Time date = new java.sql.Time(df.parse(value).getTime());
		return date;
	}
}

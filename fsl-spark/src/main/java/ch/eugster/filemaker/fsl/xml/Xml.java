package ch.eugster.filemaker.fsl.xml;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.InputStream;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.commons.io.FileUtils;
import org.json.JSONObject;
import org.json.XML;
import org.xml.sax.helpers.DefaultHandler;

import com.fasterxml.jackson.databind.JsonNode;

import ch.eugster.filemaker.fsl.Executor;

public class Xml extends Executor
{
//	private static Logger logger = LoggerFactory.getLogger(Camt.class);

	public static final String IDENTIFIER_KEY = "identifier";
	
	public String convert(String request)
	{
		if (createRequestNode(request))
		{
			doConvert();
		}
		return getResponse();
	}
	
	private boolean doConvert()
	{
		boolean result = true;
		JsonNode xmlFileNode = getRequestNode().findPath(Key.XML_FILE.key());
		if (xmlFileNode.isTextual())
		{
			File file = new File(xmlFileNode.asText());
			if (file.isFile())
			{
				result = readXmlFileAndConvertToJson(file);
			}
			else
			{
				result = addErrorMessage("'" + file.getName()+ "' is not a valid xml file");
			}
		}
		else if (xmlFileNode.isMissingNode())
		{
			JsonNode jsonFileNode = getRequestNode().findPath(Key.JSON_FILE.key());
			if (jsonFileNode.isTextual())
			{
				File file = new File(jsonFileNode.asText());
				if (file.isFile())
				{
					result = readJsonFileAndConvertToXml(file);
				}
				else
				{
					result = addErrorMessage("'" + file.getName()+ "' is not a valid json file");
				}
			}
			else if (jsonFileNode.isMissingNode())
			{
				JsonNode xmlContentNode = getRequestNode().findPath(Key.XML_CONTENT.key());
				if (xmlContentNode.isTextual())
				{
					String xml = xmlContentNode.asText();
					result = convertXmlToJson(xml);
				}
				else if (xmlContentNode.isMissingNode())
				{
					JsonNode jsonContentNode = getRequestNode().findPath(Key.JSON_CONTENT.key());
					if (jsonContentNode.isTextual())
					{
						String json = jsonContentNode.asText();
						result = convertJsonToXml(json);
					}
					else if (jsonContentNode.isMissingNode())
					{
						result = addErrorMessage("missing argument, one of '" + Key.XML_FILE.key() + "', '" + Key.XML_CONTENT.key() + "', " + Key.JSON_FILE.key() + "', or '" + Key.JSON_CONTENT.key() + "'");
					}
				}
			}
		}
		return result;
	}

	private boolean readXmlFileAndConvertToJson(File file)
	{
		boolean result = true;
		try
		{
			String xml = FileUtils.readFileToString(file, "UTF-8");
			result = convertXmlToJson(xml);
		}
		catch (Exception e)
		{
			result = addErrorMessage("'" + file.getName()+ "' is not a valid xml file");
		}
		return result;
	}
	
	private boolean readJsonFileAndConvertToXml(File file)
	{
		boolean result = true;
		try
		{
			String json = FileUtils.readFileToString(file, "UTF-8");
			result = convertJsonToXml(json);
		}
		catch (Exception e)
		{
			result = addErrorMessage("'" + file.getName()+ "' is not a valid json file");
		}
		return result;
	}
	
	private boolean convertXmlToJson(String xml)
	{
		boolean result = true;
		try
		{
		    SAXParser saxParser = SAXParserFactory.newInstance().newSAXParser();
		    InputStream stream = new ByteArrayInputStream(xml.getBytes("UTF-8"));
		    saxParser.parse(stream, new DefaultHandler());
			JSONObject jsonObject = XML.toJSONObject(xml);
			getResponseNode().put(Executor.RESULT, jsonObject.toString());
		}
		catch (Exception e)
		{
			addErrorMessage("xml content is not valid");
		}
		return result;
	}
	
	private boolean convertJsonToXml(String json)
	{
		boolean result = true;
		try
		{
			JSONObject jsonObject = new JSONObject(json);
		    getResponseNode().put(Executor.RESULT, XML.toString(jsonObject));
		}
		catch (Exception e)
		{
			result = addErrorMessage("json content is not valid");
		}
		return result;
	}

	public enum Key
	{
		// @formatter:off
		XML_FILE("xml_file"),
		XML_CONTENT("xml_content"),
		JSON_FILE("json_file"),
		JSON_CONTENT("json_content");
		// @formatter:on

		private String key;

		private Key(String key)
		{
			this.key = key;
		}

		public String key()
		{
			return this.key;
		}
	}
}

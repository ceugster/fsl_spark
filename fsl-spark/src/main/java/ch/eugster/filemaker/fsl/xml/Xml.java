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
import com.fasterxml.jackson.databind.node.ObjectNode;

import ch.eugster.filemaker.fsl.Executor;

public class Xml extends Executor
{
//	private static Logger logger = LoggerFactory.getLogger(Camt.class);

	public void convert(ObjectNode requestNode, ObjectNode responseNode)
	{
		JsonNode xmlFileNode = requestNode.findPath(Key.XML_FILE.key());
		if (xmlFileNode.isTextual())
		{
			File file = new File(xmlFileNode.asText());
			if (file.isFile())
			{
				readXmlFileAndConvertToJson(file);
			}
			else
			{
				addErrorMessage("'" + file.getName()+ "' is not a valid xml file");
			}
		}
		else if (xmlFileNode.isMissingNode())
		{
			JsonNode jsonFileNode = requestNode.findPath(Key.JSON_FILE.key());
			if (jsonFileNode.isTextual())
			{
				File file = new File(jsonFileNode.asText());
				if (file.isFile())
				{
					readJsonFileAndConvertToXml(file);
				}
				else
				{
					addErrorMessage("'" + file.getName()+ "' is not a valid json file");
				}
			}
			else if (jsonFileNode.isMissingNode())
			{
				JsonNode xmlContentNode = requestNode.findPath(Key.XML_CONTENT.key());
				if (xmlContentNode.isTextual())
				{
					String xml = xmlContentNode.asText();
					convertXmlToJson(xml);
				}
				else if (xmlContentNode.isMissingNode())
				{
					JsonNode jsonContentNode = requestNode.findPath(Key.JSON_CONTENT.key());
					if (jsonContentNode.isTextual())
					{
						String json = jsonContentNode.asText();
						convertJsonToXml(json);
					}
					else if (jsonContentNode.isMissingNode())
					{
						addErrorMessage("missing argument, one of '" + Key.XML_FILE.key() + "', '" + Key.XML_CONTENT.key() + "', " + Key.JSON_FILE.key() + "', or '" + Key.JSON_CONTENT.key() + "'");
					}
				}
			}
		}
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
			responseNode.put(Executor.RESULT, jsonObject.toString());
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
		    responseNode.put(Executor.RESULT, XML.toString(jsonObject));
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

//	public void convertWordToPdf(Object[] parameters)
//	{
//		InputStream source = new File(String.valueOf(parameters[0])); 
//		OutputStream target = new File(String.valueOf(parameters[1]));
//		IConverter converter = LocalConverter.builder().build();
//		if (converter
//				.convert(source).as(DocumentType.MS_WORD)
//		        .to(target).as(DocumentType.PDF)
//		        .execute())
//		{
//			
//		}
//	}
}

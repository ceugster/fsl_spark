	package ch.eugster.filemaker.fsl.camt;

	import java.io.File;
import java.util.Objects;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.prowidesoftware.swift.model.mx.AbstractMX;
import com.prowidesoftware.swift.utils.Lib;

import ch.eugster.filemaker.fsl.Executor;

	public class Camt extends Executor
	{
//		private Logger logger = LoggerFactory.getLogger(Camt.class);

		public static final String IDENTIFIER_KEY = "identifier";
		
		private final ObjectMapper mapper = new ObjectMapper();

		public String convertCamt(String request)
		{
			if (createRequestNode(request))
			{
				doConvertCamt();
			}
			return getResponse();
		}
		
		private boolean doConvertCamt()
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

//		public String extract(String request) throws JsonMappingException, JsonProcessingException
//		{
//			if (createRequestNode(request))
//			{
//				doExtract();
//			}
//			return getResponse();
//		}
	//	
//		public String extractTags(String request)
//		{
//			if (createRequestNode(request))
//			{
//				doExtract();
//			}
//			return getResponse();
//		}

//		private boolean doExtract()
//		{
//			boolean result = true;
//			parse(getRequestNode().toString());
//			JsonNode resultNode = getResponseNode().findPath(Executor.RESULT);
//			if (!resultNode.isMissingNode())
//			{
//				String content = resultNode.asText();
//				try
//				{
//					ObjectMapper mapper = new ObjectMapper();
//					JsonNode json = mapper.readTree(content);
//					Iterator<JsonNode> ntfctns = json.get("bkToCstmrDbtCdtNtfctn").get("ntfctn").elements();
//					while (ntfctns.hasNext())
//					{
//						JsonNode ntfctn = ntfctns.next();
//						Iterator<JsonNode> ntries = ntfctn.get("ntry").elements();
//						while (ntries.hasNext())
//						{
//							JsonNode ntry = ntries.next();
//							Iterator<JsonNode> ntryDtls = ntry.get("ntryDtls").elements();
//							while (ntryDtls.hasNext())
//							{
//								JsonNode ntryDtl = ntryDtls.next();
//								Iterator<JsonNode> txDtls = ntryDtl.get("txDtls").elements();
//								while (txDtls.hasNext())
//								{
//									JsonNode txDtl = txDtls.next();
//									JsonNode prtry = txDtl.get("refs").get("prtry").get(0);
//									JsonNode amt = txDtl.get("amt");
//									JsonNode val = amt.get("value");
//									JsonNode ccy = amt.get("ccy");
//									JsonNode rltdPties = txDtl.get("rltdPties");
//									JsonNode dbtr = rltdPties.get("dbtr");
//									JsonNode nm = dbtr.get("nm");
//									JsonNode pstlAdr = dbtr.get("pstlAdr");
//									JsonNode strtNm = pstlAdr.get("strtNm");
//									JsonNode bldgNb = pstlAdr.get("bldgNb");
//									JsonNode pstCd = pstlAdr.get("pstCd");
//									JsonNode twnNm = pstlAdr.get("twnNm");
//									JsonNode ctry = pstlAdr.get("ctry");
//									JsonNode iban = rltdPties.get("dbtrAcct").get("id").get("iban");
//									JsonNode rmtInf = txDtl.get("rmtInf");
//									Iterator<JsonNode> strds = rmtInf.get("strd").elements();
//									while (strds.hasNext())
//									{
//										JsonNode strd = strds.next();
//										Iterator<JsonNode> addtlRmtInfs = strd.get("addtlRmtInf").elements();
//										while (addtlRmtInfs.hasNext())
//										{
//											JsonNode addtlRmtInf = addtlRmtInfs.next();
//										}
//									}
//								}
//							}
//						}
//					}
//				}
//				catch (Exception e)
//				{
//					result = addErrorMessage("test");
//				}
//			}
//			return result;
//		}

//		private boolean doExtractTags()
//		{
//			boolean result = true;
//			JsonNode pathNode = getRequestNode().findPath(Camt.Parameter.XML_FILE.key());
//			if (pathNode.isTextual())
//			{
//				String path = pathNode.asText();
//		        try { 
//		            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
//		            DocumentBuilder builder = factory.newDocumentBuilder();
//		            Document doc = builder.parse (new File(path)); 
//		            NodeList list = doc.getElementsByTagName("xs:element"); 
//		            for(int i = 0 ; i < list.getLength(); i++)
//		            {
//		                Node first = (Node)list.item(i);
//		                System.out.println(first);
////		                if(first.)
////		                {
////		                    String nm = first.getAttribute("name"); 
////		                    System.out.println(nm); 
////		                    String nm1 = first.getAttribute("type"); 
////		                    System.out.println(nm1); 
////		                }
//		            }
//		        } 
//		        catch (ParserConfigurationException e) 
//		        {
//		            e.printStackTrace();
//		        }
//		        catch (SAXException e) 
//		        { 
//		            e.printStackTrace();
//		        }
//		        catch (IOException ed) 
//		        {
//		            ed.printStackTrace();
//		        }
//			}
//			else
//			{
//				result = addErrorMessage("invalid argument '" + Camt.Parameter.XML_FILE.key() + "'");
//			}
//			return result;
//		}

		private boolean readXmlFileAndConvertToJson(File file)
		{
			boolean result = true;
			try
			{
				String xml = Lib.readFile(file);
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
				String json = Lib.readFile(file);
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
				AbstractMX mx = AbstractMX.parse(xml);
				if (Objects.nonNull(mx))
				{
					String json  = mx.toJson();
					String identifier = extractIdentifier(json);
					getResponseNode().put(Executor.RESULT, json);
					getResponseNode().put(Camt.IDENTIFIER_KEY, identifier);
				}
				else
				{
					result = addErrorMessage("xml content is not valid");
				}
			}
			catch (Exception e)
			{
				result = addErrorMessage("xml content is not valid");
			}
			return result;
		}
		
		private boolean convertJsonToXml(String json)
		{
			boolean result = true;
			try
			{
				AbstractMX mx = AbstractMX.fromJson(json);
				if (Objects.nonNull(mx))
				{
					String xml = mx.document();
					String identifier = extractIdentifier(json);
					getResponseNode().put(Executor.RESULT, xml);
					getResponseNode().put(Camt.IDENTIFIER_KEY, identifier);
				}
				else
				{
					result = addErrorMessage("json content is not valid");
				}
			}
			catch (Exception e)
			{
				result = addErrorMessage("json content is not valid");
			}
			return result;
		}

		private String extractIdentifier(String json)
		{
			String identifier = null;
			try
			{
				JsonNode jsonNode = mapper.readTree(json);
				JsonNode identifierNode = jsonNode.findPath(Camt.IDENTIFIER_KEY);
				identifier = identifierNode.asText();
			}
			catch (JsonMappingException e)
			{
				addErrorMessage("could not extract camt identifier");
			}
			catch (JsonProcessingException e)
			{
				addErrorMessage("could not extract camt identifier");
			}
			return identifier;
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

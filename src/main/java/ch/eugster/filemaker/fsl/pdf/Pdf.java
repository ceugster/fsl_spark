package ch.eugster.filemaker.fsl.pdf;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Base64;
import java.util.Calendar;
import java.util.Objects;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentInformation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.json.JsonMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import ch.eugster.filemaker.fsl.Executor;

public class Pdf extends Executor
{
	private PDDocument document;
	
	public void getDocumentInfo()
	{
		String content = null;
		JsonNode node = getRequestNode().findPath("content");
		if (node.isTextual())
		{
			content = node.asText();
		}
		else if (node.isMissingNode())
		{
			node = getRequestNode().findPath("file");
			if (node.isTextual())
			{
				File file = new File(String.class.cast(node.asText()));
				if (file.isFile())
				{
					if (file.canRead())
					{
						InputStream is = null;
						try
						{
							is = new FileInputStream(file);
							byte[] bytes = is.readAllBytes();
							content = Base64.getEncoder().encodeToString(bytes);
						}
						catch (Exception e)
						{
							addErrorMessage("reading file failed");
						}
						finally
						{
							try
							{
								is.close();
							}
							catch (IOException e)
							{
							}
						}
					}
					else
					{
						addErrorMessage("file not readable'" + node.asText() + "'");
					}
				}
				else
				{
					addErrorMessage("not a file '" + node.asText() + "'");
				}
			}
		}
		else
		{
			addErrorMessage("missing_paramenter 'content'");
		}
		if (Objects.nonNull(content))
		{
			try
			{
				document = PDDocument.load(Base64.getDecoder().decode(content));
				PDDocumentInformation info = document.getDocumentInformation();
				if (Objects.nonNull(info)) 
				{
					JsonMapper mapper = new JsonMapper();
					ObjectNode metadata = mapper.createObjectNode();
					String value = info.getAuthor();
					metadata.put("author", Objects.nonNull(value) ? value : "");
					value = info.getCreator();
					metadata.put("creator", Objects.nonNull(value) ? value : "");
					value = info.getKeywords();
					metadata.put("keywords", Objects.nonNull(value) ? value : "");
					value = info.getProducer();
					metadata.put("producer", Objects.nonNull(value) ? value : "");
					value = info.getSubject();
					metadata.put("subject", Objects.nonNull(value) ? value : "");
					value = info.getTitle();
					metadata.put("title", Objects.nonNull(value) ? value : "");
					SimpleDateFormat formatter = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
					Calendar calendar= info.getCreationDate();
					metadata.put("creationDate", Objects.nonNull(calendar) ? formatter.format(calendar.getTime()): "");
					calendar= info.getModificationDate();
					metadata.put("modificationDate", Objects.nonNull(calendar) ? formatter.format(calendar.getTime()): "");
					getResponseNode().put(Executor.RESULT, metadata.toString());
				}
			}
			catch (Exception e) 
			{
				addErrorMessage(e.getLocalizedMessage());
			}
		}
	}

}

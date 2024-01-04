package ch.eugster.filemaker.fsl.qrbill;

import java.util.Objects;

import com.fasterxml.jackson.annotation.JsonProperty;

import net.codecrete.qrbill.generator.GraphicsFormat;
import net.codecrete.qrbill.generator.Language;
import net.codecrete.qrbill.generator.OutputSize;

public class Form
{
	@JsonProperty("graphics_format")
	private GraphicsFormat graphicsFormat;

	@JsonProperty("output_size")
	private OutputSize outputSize;

	@JsonProperty("language")
	private Language language;

	public GraphicsFormat getGraphicsFormat()
	{
		return Objects.isNull(graphicsFormat) ? GraphicsFormat.PDF : graphicsFormat;
	}

	public void setGraphicsFormat(GraphicsFormat graphicsFormat)
	{
		this.graphicsFormat = graphicsFormat;
	}

	public OutputSize getOutputSize()
	{
		return Objects.isNull(outputSize) ? OutputSize.A4_PORTRAIT_SHEET : outputSize;
	}

	public void setOutputSize(OutputSize outputSize)
	{
		this.outputSize = outputSize;
	}

	public Language getLanguage()
	{
		return Objects.isNull(language) ? Language.DE : language;
	}

	public void setLanguage(Language language)
	{
		this.language = language;
	}
}

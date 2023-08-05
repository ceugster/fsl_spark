package ch.eugster.filemaker.fsl.qrbill;

import java.util.List;
import java.util.Objects;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import ch.eugster.filemaker.fsl.Executor;
import net.codecrete.qrbill.generator.Address;
import net.codecrete.qrbill.generator.Bill;
import net.codecrete.qrbill.generator.BillFormat;
import net.codecrete.qrbill.generator.GraphicsFormat;
import net.codecrete.qrbill.generator.Language;
import net.codecrete.qrbill.generator.OutputSize;
import net.codecrete.qrbill.generator.ValidationMessage;
import net.codecrete.qrbill.generator.ValidationResult;

/**
 * Generates swiss qrbills from json parameters. Based on the works of
 * net.codecrete.qrbill/qrbill-generatory by manuelbl,
 * 
 * @author christian
 *
 */
public class QRBill extends Executor
{
//	private ObjectMapper mapper = new ObjectMapper();
//
//	private Parameters parameters = loadDefaultParameters();

	public void generate(ObjectNode requestNode, ObjectNode responseNode)
	{
		try 
		{
			Bill bill = new Bill();
			bill.setAccount(checkString(requestNode, Key.IBAN.key()));
			bill.setReference(checkString(requestNode, Key.REFERENCE.key()));
			bill.setAmountFromDouble(checkDouble(requestNode, Key.AMOUNT.key()));
			bill.setCurrency(checkString(requestNode, Key.CURRENCY.key()));
			bill.setUnstructuredMessage(checkString(requestNode, Key.MESSAGE.key()));

			JsonNode creditor = requestNode.get(Key.CREDITOR.key());
			if (JsonNode.class.isInstance(creditor))
			{
				Address address = new Address();
				address.setName(checkString(creditor, Key.NAME.key()));
				address.setAddressLine1(checkString(creditor, Key.ADDRESS_LINE_1.key()));
				address.setAddressLine2(checkString(creditor, Key.ADDRESS_LINE_2.key()));
				address.setCountryCode(checkString(creditor, Key.COUNTRY.key()));
				bill.setCreditor(address);
			}

			JsonNode debtor = requestNode.get(Key.DEBTOR.key());
			if (JsonNode.class.isInstance(debtor))
			{
				Address address = new Address();
				address.setName(checkString(debtor, Key.NAME.key()));
				address.setAddressLine1(checkString(debtor, Key.ADDRESS_LINE_1.key()));
				address.setAddressLine2(checkString(debtor, Key.ADDRESS_LINE_2.key()));
				address.setCountryCode(checkString(debtor, Key.COUNTRY.key()));
				bill.setDebtor(address);
			}

			JsonNode form = requestNode.get(Key.FORMAT.key());
			if (JsonNode.class.isInstance(form))
			{
				BillFormat format = new BillFormat();
				format.setGraphicsFormat(checkGraphicsFormat(form));
				format.setLanguage(checkLanguage(form));
				format.setOutputSize(checkOutputSize(form));
				bill.setFormat(format);
			}

			ValidationResult validation = net.codecrete.qrbill.generator.QRBill.validate(bill);
			if (validation.isValid())
			{
				byte[] swissqrbill = net.codecrete.qrbill.generator.QRBill.generate(bill);
				responseNode.put(Executor.RESULT, swissqrbill);
			}
			else
			{
				List<ValidationMessage> msgs = validation.getValidationMessages();
				if (!msgs.isEmpty())
				{
					for (ValidationMessage msg : msgs)
					{
						addErrorMessage(msg.getMessageKey() + ": '" + msg.getField() + "'");
					}
				}
			}
		} 
		catch (Exception e) 
		{
			addErrorMessage("invalid_json_format_parameter '" + e.getLocalizedMessage() + "'");
		}
	}
	
	private String checkString(JsonNode requestNode, String key)
	{
		if (Objects.nonNull(requestNode))
		{
			JsonNode node = requestNode.get(key);
			return Objects.nonNull(node) ? node.asText() : null;
		}
		return null;
	}

	private Double checkDouble(JsonNode requestNode, String key)
	{
		if (Objects.nonNull(requestNode))
		{
			JsonNode node = requestNode.get(key);
			return Objects.nonNull(node) ? node.asDouble() : null;
		}
		return null;
	}
	
	private GraphicsFormat checkGraphicsFormat(JsonNode format)
	{
		JsonNode f = format.get(Key.GRAPHICS_FORMAT.key());
		if (f == null)
			return GraphicsFormat.PDF;
		return GraphicsFormat.valueOf(f.asText());
	}

	private OutputSize checkOutputSize(JsonNode format)
	{
		JsonNode f = format.get(Key.OUTPUT_SIZE.key());
		if (f == null)
			return OutputSize.QR_BILL_EXTRA_SPACE;
		return OutputSize.valueOf(f.asText());
	}

	private Language checkLanguage(JsonNode format)
	{
		JsonNode f = format.get(Key.LANGUAGE.key());
		if (f == null)
			return Language.DE;
		return Language.valueOf(f.asText());
	}
	
//	private Parameters loadDefaultParameters()
//	{
//		Parameters params = null;
//		if (Objects.isNull(params))
//		{
//			Path cfg = Paths.get(System.getProperty("user.home"), ".fsl", "parameters.json");
//			File file = cfg.toFile();
//			if (file.isFile() && file.canRead())
//			{
//				try
//				{
//					params = mapper.readValue(file, Parameters.class);
//				}
//				catch (Exception e)
//				{
//					params = new Parameters();
//				}
//			}
//		}
//		return params;
//	}
//	
	public enum Key
	{
		// @formatter:off
		IBAN("iban"), REFERENCE("reference"), AMOUNT("amount"), CURRENCY("currency"), MESSAGE("message"),
		CREDITOR("creditor"), DEBTOR("debtor"), NAME("name"), ADDRESS_LINE_1("address_line_1"), ADDRESS_LINE_2("address_line_2"), COUNTRY("country"),
		FORMAT("format"), GRAPHICS_FORMAT("graphics_format"), OUTPUT_SIZE("output_size"), LANGUAGE("language");
		
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
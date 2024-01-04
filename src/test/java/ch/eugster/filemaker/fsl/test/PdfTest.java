package ch.eugster.filemaker.fsl.test;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Base64;
import java.util.Objects;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeoutException;

import org.eclipse.jetty.client.api.ContentResponse;
import org.eclipse.jetty.client.util.StringContentProvider;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import ch.eugster.filemaker.fsl.Executor;
import ch.eugster.filemaker.fsl.Fsl;
import ch.eugster.filemaker.fsl.pdf.Pdf;

public class PdfTest extends AbstractTest
{
	private ObjectMapper mapper = new ObjectMapper();

	private static String sourceFilename = "src/test/resources/pdf/document.pdf";

	protected Pdf pdf;
	
	@BeforeEach
	protected void beforeEach() throws Exception
	{
		if (Objects.isNull(pdf))
		{
			Fsl fsl = Fsl.getFsl(System.getProperty("user.name"));
			pdf = Pdf.class.cast(fsl.getExecutor("Pdf"));
		}
	}
	
	@AfterEach
	protected void afterEach() throws Exception
	{
		if (Objects.nonNull(pdf))
		{
			Fsl fsl = Fsl.getFsl(System.getProperty("user.name"));
			pdf = Pdf.class.cast(fsl.getExecutor("Pdf"));
		}
	}
	
	@Test
	public void testDocumentInfo() throws IOException, InterruptedException, TimeoutException, ExecutionException
	{
		File file = new File(sourceFilename);
		InputStream is = null;
		try
		{
			is = new FileInputStream(file);
			byte[] content = is.readAllBytes();
			ObjectNode requestNode = mapper.createObjectNode();
			requestNode.put("content", Base64.getEncoder().encodeToString(content));

			ContentResponse response = client.POST("http://localhost:4567/fsl/Pdf.getDocumentInfo").
					header("Content-Type", "application/json").
					content(new StringContentProvider(requestNode.toString())).
					accept("application/json").
					send();
			
			JsonNode responseNode = mapper.readTree(response.getContentAsString());
			assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
			JsonNode result = mapper.readTree(responseNode.get(Executor.RESULT).asText());
			assertEquals("SwissQRBill", result.get("author").asText());
			assertEquals("SwissQRBill", result.get("creator").asText());
			assertEquals("", result.get("keywords").asText());
			assertEquals("SwissQRBill", result.get("producer").asText());
			assertEquals("", result.get("subject").asText());
			assertEquals("", result.get("title").asText());
			assertEquals("2023/03/29 10:57:18", result.get("creationDate").asText());
			assertEquals("", result.get("modificationDate").asText());
		}
		finally
		{
			if (Objects.nonNull(is))
				is.close();
		}
	}

	@Test
	public void testDocumentInfoWithInvalidContent() throws IOException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put("content", "Das ist ein normaler Text, der sicher nicht gelesen werden kann.");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Pdf.getDocumentInfo").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("Illegal base64 character 20", responseNode.get(Executor.ERRORS).get(0).asText());
	}

	@Test
	public void testDocumentInfoWithInvalidFileParameter() throws IOException, InterruptedException, TimeoutException, ExecutionException
	{
		File file = new File(sourceFilename);
		InputStream is = null;
		try
		{
			is = new FileInputStream(file);
			byte[] content = is.readAllBytes();
			ObjectNode requestNode = mapper.createObjectNode();
			requestNode.put("file", Base64.getEncoder().encodeToString(content));

			ContentResponse response = client.POST("http://localhost:4567/fsl/Pdf.getDocumentInfo").
					header("Content-Type", "application/json").
					content(new StringContentProvider(requestNode.toString())).
					accept("application/json").
					send();
			
			JsonNode responseNode = mapper.readTree(response.getContentAsString());
			assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
			assertEquals(1, responseNode.get(Executor.ERRORS).size());
			System.out.println(responseNode.get(Executor.ERRORS).get(0).asText());
			assertEquals("not a file 'JVBERi0xLjMKJf////8KOCAwIG9iago8PAovVHlwZSAvRXh0R1N0YXRlCi9DQSAxCj4+CmVuZG9iago3IDAgb2JqCjw8Ci9UeXBlIC9QYWdlCi9QYXJlbnQgMSAwIFIKL01lZGlhQm94IFswIDAgNTk1LjI3NjUgMjk3LjYzODI1XQovQ29udGVudHMgNSAwIFIKL1Jlc291cmNlcyA2IDAgUgo+PgplbmRvYmoKNiAwIG9iago8PAovUHJvY1NldCBbL1BERiAvVGV4dCAvSW1hZ2VCIC9JbWFnZUMgL0ltYWdlSV0KL0V4dEdTdGF0ZSA8PAovR3MxIDggMCBSCj4+Ci9Gb250IDw8Ci9GMiA5IDAgUgovRjEgMTAgMCBSCj4+Cj4+CmVuZG9iago1IDAgb2JqCjw8Ci9MZW5ndGggODUwOQovRmlsdGVyIC9GbGF0ZURlY29kZQo+PgpzdHJlYW0KeJyl3c+OG8fZxeE9r4I34HZX9X9A8MJOYiCLAIm1M7ywR6M4gB0gn4Hk9r+3m+R7DjWa+tEeCFE01COyp9l16lQ1JZdzHz++KPFT3ZZuHtY6nZ9+PZVl6pZxHeJx+0Lkl1PfLdP5f6cvv/2tnP/52+n7ci4/hP5w+vJPz//919PzP779+vzNd6f+eIHvvvnb6bt4nqWbpjKch7Gbxul45rUrw3w80s/n48uyfzl0ZZ3O+x9Y5zG+rt00lPNT/IG5G+oUh7J243Ec+yPbsh/c0tUw+9fLFi9S566fyvEcpd/y66fjMIZpvRNDPxxf12U7vt6myzPUOMyn4zCn5SpKDbF18Qfi66mb4nD339+2Gl+P8Sevf2Db5uOBYbh8Y9Nx2EO3XL6tZTm+mq8+Dr6Ol0fG7fgm5rIcf75sl6+nZTy+Xsf1eiLGuRyHMPcXsdV6HOI27a8R79m6HadlWrfjT8TB1f54pJT1fHx9vEZ8U2u9fH19xqnOtz+xXY5iPY4i/v92VON0fF37chz1Ol7+RBx/6Y9HpuMoarfM21UcX07rejzB1I/HH6jdMNbjRePpDjFPlzPdb/vXQ7zP1zM/jdeD6q+HPSzL5SCOUz3FNz4fX0/lciL2b/PyJ+Zlur6fF7BezsN+8PuZinN8XFLDVm/fxVCOy24d1uMii8PYnyIuzOtB7c+wX7Xr5RD6oV6u6/VykHH2j5fc4uIcj2e4v/Z/Of18uaz6yxkftnIbD2N/+faH9XLxTtvlmzsGRLy78+Xr/QAv18EyXd6Tpb8MmeEYq0O828tx5fX95ZSXWm4X87gdj9RhOsRldMe7Pa7nT4/q6TjSsRu3y9s0lu040vjT4/0bVfr+eqTH+7iNl3elbsP1nV6u7/06LMcf6Mt0ObB1Pr6epnq9uK5vW+0v4DJM4zvcrmPgev36Me3Hadnz9Ns1e357+vfp46kcv778HPn2n+sDL5Pv6/dpSxzJEmkTV8o+zpfz+19PX/4lRko5v/94+v7dOM0fln6e5zI/zzGel0iyr86lP7+b13mat/n5q3P/w/n9X09/fn/6++99zSmu7zLGpXJ91fn6oj99dY537d38cX5exvlj7Wv8b/oxDmKdn+YasVFrvx/SW158jKtqWG8vXs7r9cX3Syt+qnv6lWEb9kDZ4lcxRONXlx9r/GqNR/uhxK/qWw5jiCTe+heHEfH0Mb7TaYkn/2KM872M+5UyD3EKno8zP+6/WfuxxMmK3x+XtxxFDOI61s8cxTLNU5zrSIZ4qac45/tJ2E9GPN+bXrCPK3t88YLxxNPlrI41vsdpfnrLi/QREMPL6yvO7TTP18s4TmJ89bz8+IYXKtu4T7ef+25KXjPxw359OY1DXEPjsBy/jqRJc1xZw1ve0WN61Tuq7/3TURTv7H6N1f1c5NB+ywvPMefM88tLqY/3s9QP44f9heOiLvsI3x/q96tsGZfLxf3hiJrn+uF2sccxhVyuF/xbjmyK2a98ZsQvcQl8PLJtv8jvju+nyJ9lrymX34w39E3nZtyi47wc7MdVcLz3+7Bf6vVUlP3nt7xc3++Vqn9xDcSFtUfK8zivyz7CI9p//8vsfW2YptdfZtzja4xzW66X1Vuu562P3x/LZ8J6nN9w7J9/2mPUXrL/OX6e/sBxlzX65N5W5r5b13V7cXKOq/wYix/iNA1xnUXYxY8/8mJrVLBapldn8dugP15k+yRT9afHaGbrFsdcY8WyzP6L/3u2V5n6WKDM88SwRoGPywBdtNc+GhI/4T5nb+OCbt1XH/2CTziXaGgxDNBFx1u3eB8JLiU6ao0xjnBfLsSY52eMtrnEco+fcV9ZRgHHZ1yjn9d1GPAZt+ivy3qUfIBbzOZl4vNY+tLN9bImIxnfdz8OK8sSC6yylI1ffV9Ar/GkFeUwRYXv1+kBGQVujlkFL8oy1n09EWWa5RpLxHlcC8opvqOYWfsH5NbtXajnV5/3Fe62lAfkdKzM+CRFKVi3WMrggCxLvy80+ZLbF1uB441iGe2sXxa83I8Nh3EaxteecquxTNseiEpBCBiDSzfWJToxyygR89Rvr41zybXv4kyWys+5xhDa+qU+8JzRK+dYr7OMuXUbS1zLLCFnTEYOx/U58XOWKP/LHGvWRyhEjVHKGqexeJ/GtX9tdDiNPrLVsX8tQ4zGSm0saz88QIfSRceJCfgBCnHnFPLO6BgDb43QfS2cnMYZKOu6vJY5TiEcjUZTWbaxXx94VgpSp5CkRilKjWL47PuksRjB7Lk5aF9y0L4E2+0rHWVeOihVgpSNN0jtKyHmXULoSoJROfthGl/NkBvEsEsIpSph6deuHyOTWGLQSK7xZvd9eTURUmIipRz2mwFD5SuIQybl2Heln/v51XErCfXLJOWWJGVRymnqIl7qxq+OSSRJQZQSc0gSmpokNDWD0NRMUgbOXQzFYeYClpBSUJBiUBJyMCEFoSAlYUpMOElYYEpu+83C4bjBRJIiKSXXL6NUv4xS0omWCIZlO+4sIaVUFK19tx47gkwx7UQxxIxSihmlGDNK9UsUg8wo5ZMoBpRRSihRiiiTlFFGKaTWbowVwoIZdXMUUekooRJCQN0c5VM6iqeEVNRukDa1EmL/Skjr0hvE/pWQ1poJKRVvkENRkoIuJeacJK1dJWnpKknRKUmFMmWNd6ifan1ATl0E4lr51WnnTxJXzZIU8JJUZ1PiTCBJ1TMlroIla7cNZd4eeE6aBSSpzqbEOUCSpoCUNAMI0gQg2c7/2vf7x5WmhUqqwXYQO2wnsct2FJuELHbZvsNgEtLYJJRPl+2YNUl3Ge5ou3ze0XYmO4VYvKPtXHQK5dMppNgdbYeTUyifd7RdPu9ou3w6hYC6o+2EcgrB4xQHf+2WyIkVx/7Nwc1Ug82WaK7dEg02W6K59s1UQUqxdBRiCdt3OgRhrWuQEuwGYYPQICXdDcLtEIPt3mmQAvEGqXe6pIxLCW3SJYVhSgw4yXbzc9nenjQJfc4lRWZKTExJSsGU0OdcUgZKtpfqJjEsJZt3hx2265zJdp1z2K5zLpt3hx1S9I/duo7LOnJupYTtPJfU0lLCXV+X7cW1ScwkSepzKTlrRDFCjFKfEsW4MdpekzrFGBHFdDBKfcooJYnR9iLSKaaO0faS745S7ixdX+sQXYH6jyAVIEloQIJUgVJSBxIcumizdX117EnSSi4l1qCUmBGS1G9SYkZIUkZIUsVJyWkiCltmd5RajtH2ptkdpTgTxYwySp1IFJeHRinOjFItEoV9rjtKyWeUkk8U48woLTqNUuEy+nDylQiVca2lPHAA1LlMUukSpdZlkmqXUapT2/73wraIPVpKJ8S1tCTNEZIwRwjSHJESo1+SVsqSNEmkxMopSdOJJC2rU+K6WrL9gRqTcEfHJU08KXnJbJS2EI3SJCWKM49Rmk6MUjsWxZnHaPvuilOcTkSxHRul4DdKlVcU5wijNEcYpTnCKBVpUVyWG4U5wiTNEaI0R5iEBbdJmCJK6eq2BKYpQpCmCJPt9YbB9gwhCGsDgzBBmISdBkmKfZMQ5pIU5iYhzCUpzCUx9ozS8t0pZJlRatFOIfactj+L7ZRatFNo0UYpy5xCljmFLHMK1dgoVWOn7dgzCVnmsp1lLinLhm4c9r9/isEjCCVWEgNFEhqnSYqelNQjTVL0pKRtCZMUPSkxeiShR0rirWinUA6NUjl0CtsSTilQjUKPNEo90incqzGK2WuUstcoZa9R2OwwSu3UKbRToxjTRimmjVJMG6WYFqV26hRi2iS0U6OU0yYpp6duWqIgcOdMiJ1TEqJfkKJfEtppQmqngjSZpMTgl6TgT0k71yYp+FNi8KfkODcKOwhOKflFae/aKcW5UYpzUazSRmFD2ijGuVGKc6MU56IYvKKYpkYpIkWxyRql4BOlDdmydEuJa483T01CSglSSqWk7UtJTIqUOP5T0kdzTFLxk4Q7V5KcKUZhq9EpZYooZopRKp5GqXgapaQSxTZplJLKKBVPUYwfUdrANEpbjU6p9xmlpDJKSSWKZU6UepdJ6l3b/s9C9uWB7iMJt2QkqfwI0vpYktbHKXFrThI+BGSSYjIlFipJ2u5LiYEqSdVLEj5eYJJCWpJCOiUnr1FKXlFMXqNU0UQxI41S8IliRTNKGWmUKpooxqlR2u0UxYwUxeATxd5lFDLSJGRkLfs/oTpXLnMmIU1NtleSBiF2JWkpaRJy1yRUSUlanpqENJWkjDQJH76UpDQ1CRlpEjLSJGSkJBZZpxCnRikjnUI7dQrt1CjFqVO4eeQUktcoJa9TiFOnsIFplO4zOYXkdQpLbqPUeY3SOtop5LlTWHIbhbtHLtvJf/x71kOtuNUoiLkv2V7EC0I1NkhhnpIKr0mKaEmoxpIYvJJwk0kSI1qSgleSgjclbjU6hXJqlFblTqGcGqVNQaeUpqKYUEahRhrFgBClUW8SqqHTRwOCq+Hc1blfhweqoSQUvoQ48lPieJak8SxJ5SwljlJJGqWSNEpTcjsxSkPPKA09UdoQc0pDzygVGVEcpUZplBqlyiGKPUIUe4QojmijMKL3/0DHOg485SfE8SxJSz1JGPmCtNRLSTVCkMJEklZ6KTFMJClMUmKNkKQaIQm3LCQxyiRhh80kRVlKXukZpW5ilBaFopilRmmlJ4oBaZRqjFHKUqOUpaK40hPFLDVKWWqUlm+idHfDKWzHOYWPqhjFMDdKKz1RzH2j0M+GvpuGOS4SCn5BCn6TEPySFKkm4faGJHU+kxDTknTTwiRsyElSpJqESJXEoDRK+1xOIf2cQpM0SttMTmGD3ynklFHKKaeQU07h03dGKXyM0tLQafsTbSZhETnUbl6GiFwMiYQYEpLtDSFBaHIGKSJS0oaQSQqTlNTPTELrksSBLwk78SZhqWkS+pkkli6nULqMYkIZpdgRpe11p1CPjGJCGaWEEsXYMUqxI0r1yCksYJ1SmBmFJmUUc0/04TTDIuW0vYAe9v9yYtlflPJMsL3YNQiLXUlaw5qE+5omKU1TYppKwmpXEtNUEla7Jil3U2LypeQ4Mwpb4UYx+YxSNzMKK1OnlKdGYUPQKC1ijWI5NEp5Kkobgk5huWkU40yUFoZOYWFolJLPJCWfUUi+ef9v264Tl8OEWA4lIUwFaakpCbErSLGbkpqpIOVzSsxSSWqmkrTMTYmLV0lKXUlK3ZTYdlNy6opilBqlKDVKUWqU8lGUdu6MYj4apWoqivlolKqpKJZIo1QijVLqimLqGqW+aRTW2SYpykUpyk0+ms+8Gbh1sRjY/9viFOUJMcolIXcTYjWVpF1DSSqxKTFOJSlOJanupsQ4laQ4TYklNiUv341SiRXF5DVKyWuUmqko3V5xSs3UKG0KiGIzFcXkNUrJa5Q2BUQxeY1S8hqFv+ZhFJNXlELSJISkSQjJsXQ1LruZO6dJqJKSUCUNQpU0CSkpSaXTJKSkScg+SSqdkpRokngfxCmEn1OonUYp/JxClzRKieYUEs0odUmn0CWN0t0Vp5BoTmFZbpRqp1NINKeQaEahIJqEmHJJMTV0wxyFkyuayfZq22C7ywnS/VKTsNFnEpqPJC0kTcJtE0lOFKMUE6IYE0ahIzmFjuQUVqdGMXyMwkLWKdQpoxg+RuEvGxjFRBHFRDEKdcoorU6dQqJIQkVyCdkzxaS71QV30ASpzAjCks8klZmUWGZSYkilpHumkhwTRikmRHFAi+IoNUqj1CiNUlFsE0ZpQIvS+sgpFQ+jNPZFsSKI0ig1CRtD49LNw7yuD8zmkjBMBWl1khJHX0pcIEjSAiEljlNJWkpIUu2QhK0Zk1RQUtJHv0xSlZGkjErJGWWUFkdGKc5EscoYpcWRKH2wwyktjkQxzoxSnBmFnSGnVGVEMfmM0pJLFFuPUWo9RmkdJUp5avLhPMXaY5KWXFsXzzOueCdUENNcEtZmgrQrJQkThCBNECmp8QnS9pUkdcOUOOdI0pyTEueclJjlKXHv3CiuC41SmBqlMDVKjVMUc9colVNRTEijlJCiWCONUkIapRopiglplBJSFJeQRmGb3SnkrknKXVFamJqEhDYJCT2VbovLdMMbnIIQkoKUU5J049Ak5JQktVNJapImoUlKYpN0CvXQKeSkUYopp7DTZZRWu0ap8zmFRDNKMeUUsscoFTmnEFNOYWHsFBLNKCVK1J1+3OYNS5cgVCmDMPYTQu0xSCEhCUttk/DZBkkMnpRUe0zCAtokLIslqUqZhDuBkhxRRimijFJEiVI/c0ppZpTSTJT6mVPoZ04pI0WpnxmlFaxT6GdO4WMQTilORemvUTil5DVKyWuU4lSUWp/Th5OXupzLdpdzSV1ujriPS79i8Cek4BeE4E+IeZ6SbkeYpJSWpHqYkjZEJbEeSsI9U5OU0ilxn9EpRa8o5qlRylNR2jw0ihFhlCLCKEWEUYoIURz3ojTu1y4acVx3OKAEYatJkjqXIGw1maTRnBJHsySN5pTYpCRpsZcSR7MkjWZJWham5NFslDqXURr4orR95pTqmVHYPnNKyWOUkkcUk8co1TNRusPhlJqcKC5hjVI9E8U8E6U9MafUjoy298RcQp7OfbSSWPJikRGkPX6TkKeSFJMmISZNwiLWJCxiJSl6TcL9AJMQ0iahcklS5TIJcW4SltCSFPwmIfhNQvBL4mLbKaS5UUpzp5C7TiF3jVLuGqUwdQr7gU4hd43SCtop5K5TWBYbpc7rFDqvU+i8TmGOMEpzhFOYI4xCPXZZ238d12l7BT3XmBr6UhZa7Rqk2USyXeMFoZwbpBlCkmaIlJjmkpTmKTHNU9LmpSTnmVFop0Zpm88p5Zko3d8wSo3PKYWUUdjmM4pxYpTiRBQzwihVPlEc+UZhR2weY3QuS+UimZAyQhBGviAVzpQ4TFPS3QiTsIaWxIIkCbtXkrjeNUprSKdUUERxQBuFhaFR7AeiOOkbpaEnikPPaHtdZpJm57nbL484lfQhLZMwnARplKSkmVSQZlJJmklT4gpKkmbSlDiYJWnOlaRhL0lrrZQ4j0vSCiolB4QolgOjVA6MUuwYpSWUKFYOo3Bn0Siui4zSukgU24lRyj2jtC4SxdwzCh+pcEqLHVFcwRiFNDVJlUf04eDlcmSUytEWOVqjwWJIC0LnSYgbYilx8yolRpokRZokbR+lxPCTpPCTpB6VkmPSKG0fiWL2iWKgGaVAM0rtzCi1M1HcaTJKizhRjEmjtDITxXpolGLSKMWkKDZJoxSTRiEmJTHRjMJGj0nIvkidOPNTXEywMBSkzmuyHacGYZvJZDt4BaEcG4SENgk1WpIqryQVWZOQ+pIYfEapSjqF0ucUlqVGKc2MUpMzShHlFPaZnEKaGaU0cwpbUk6hyRml0mWUgieWenVd40RhRiTEjJCEjBCEkZ+Q7laapKEvSUNfElbQklT4TMJa2yQFT0oMHkkokSahREpSiTQJm3Em4R6kJN6DdEpZKkprbafQN51S7IpSNXUK1dQp/BMARqlEOoUVvFFMaKOU0EZhWW4Uw9woVFOnUE2dUu6LUjV1SlOE6ONTBHVTl9RNpwjdaRm5SApSkZSE6SQhFUlBmiNS0u0NScxJSUo/SVhsS3JOimKRNErtUJSWxU4ppoxS5xSlFbRRzB5RrHyitCo1SiPPJI28JUrMMExczwRhPCXE0iVJVSol3TgwSWUmJQ6TlLgn5ZTqhChO/KI4m4viysgoXdCiOEMapRnSKE17ojhMjNIMabS9eWMSNq5d0lS6xbKg36YHNlok4Z6hJM1mgrSKkqR5T5IGdEpcG0nS2iglhoQkbbVI0oonJc7kkjSTS1JEpeSZ3CilmVH4ZL5R7AdGaXFklKqEKGakUeoHotgPjNLiSJTuLjqlZYwotg6jtDYRxeQ1SslrFJJX8vE8xYJkEgrSWmL1uk4xi1JIm2xXKYPtKmUQYl+SMlqSktckrGIkKXlNQvKahOSVpNuQJiFPJTFPnUKeOoV2aJTy1CnkqVPIU6O0NHMKeWqU8tQp3A4wSqXXKWw2OYV+7BQC3SiltFNIaaOU0k7b0WsSSu86REDMJU4l3DcQpIwUhIxMCIXXIKxLTVKYpqRyapLCNCUGX0raOpfkODMKnc8oZpRRWOw6pYwSxYwyCttHTinORGmX2yimiVHofEYxTYzCatsopokoNqm5K/FlfNsUEgmxcknCdrBJqlIpaWFsEpa7kpgTkpQTKekGn0kqXSkxUSRhESvJ2WOUskcUs8co9SOjFFNGKaZE6b6dU1jFOqXwM0rhZ5TCT5QWvE4pJ0WxoBml1mWUIlUUI9UoFTRRjFRRql0ma3tt7JRyeo013P41x68ktDlBymlJ6H0JqfcJUu9LiSktCR/DkMSUTkkbiCZpwZuSF5yi9OkGp5SSohhSohgnRmlpKIpLQ6OUPEapzInSrROn1NBEceQbhZEfF94Qw2PFhiZIEWGyHREGoaBJUkEzCQVNkoa+SShokrSQk6ShbxKGviTWLqewg+UUAsUoBYpRKkhGqUk4hZQwSinhFEqHUaoHRmnr3Cl8rMcopcRWY2UwTytP+iZhRAvSiE4Jc7lBGvqSMOubpJBISVvnJilOJKFJSNImu0noHCZhZWgSPlchiRElCdvxkpxQRmENaZTWkE4pzIzCctMprCGNYkQahdWeU1jtGaV65pQy2ihltChmtFFock4pzo3CGtIolT6n8HkZp9APjdJy0ylNPUbbtw5c0iQlCstdlw9PZ3Qr2CXV41i7zlF5uMsKthexgpjpkpS/KbFMSsIeniSXSaMU1UapTIpi/hql/DVK+SuKoWqUQlUUQ1UU488oxZ8oVlSjlD6imD5GKX2MQvpIUqSYpEiZYxEXJYF7ryBsdpmEnBCkKp2SqrQg3OQ0SQU5JdbelFhRU2KcpeQ4EaUdLKO0g+WUMkIUB74o7Uc7pS5hlEazKI0mkzSXrvGtRC/iraaEuDCVhAEqCMMuIS43JWm5mRIXkZK00yRJi8iUWCNS4riThJuGJqlwSNLSMCVXE6NUTYzSKlIUV5FGqcWI4nrPKNwzdEphZpTCzCgVHlH6/JdTikijtIoUxRpllGqUUVpFiuLS0CjFuSg2LqPUuIxC4zJJ6z1Rmk5MUjkzCus9k+05aog8iSfqe/pUm0GYo1w25yiH7RJpst34HLZ7nEmYeVy2Zx6TMPOYhA1El+1ZwiTNEne0XSOdQvTf0Xb039H2AvaOtnusU4j+O9qO/jvajn6nEP13tB39d7S91+gU8vyOtvP8jrZ3BZ1C9N/Rdj13CtF/R9vRf0eb0X8n29HvtB39d7Id/UNUuRrvUsG/+uCSAjAlLE9NQk12SWGVkhNIFANAFANAFIeqURp/ovChiTtKg0oUlsd3lAaVKF7UonRRm6SLeor1UY3JD6tHQqweklA9BKl6SDYX0g7b+1cmqc0ItlfcLikgUsLa3GV7p8skbPG7bN+Mddn+ZIlJWJu7pNYl2V7Fm+TWZbS9Nr+j7bX5HaUkFcWCZpRC12j7DoNTzGejVNBEsUqJwtL4jlKVMkpVyiilvihWKaPtVfQdpQlCFKuUUZpLRGkuMfnwXAJr4/it6I/bSv/6mxxldDqK6IRUzG4QwzQhZekNYiu7QQ4oSconSYqnlBg5KWE30CUFTkrsg5LtnTiTGCGS1BslKUBSwn1Pl5QJKXGcbzEbLmP8Po20hFTaBKm0SUJpS0jDXJDGeUosWJJUsFLiWk2SQiEl1iZJKkMpebCL4r6SUaotRmlfySjFjSimiCiGgyi2BlFcABqlsSyKuypGqQoYhV0Vkw+HCc3vJQZdzK4DpU46CB05yBzBduSkg8SRg8ARhGKRkIJJELbGE1JVSUgLr4S4mjIJXcUkdBWTEF+SlDMmYXUkSa3GJOSRJLUaSVoZmYRWYxISziS0GklqNZIUbybbkWUQEkvy0cDCe3VliKu3xuDBbEtI4SZI6SYJ8ZaQ8k2QAk6SEi4l3aqTpEZlEhqVJO2Tm6Q8TMnhZZTSS5QalVOoSUbh8xxOacvGKTQqo5hMorS54hQ2V5xC+TJKjSrOez/2MbfSsL85GnnpYK9YkMbdDWKzSEjj8wZxeN4gjqSEsIhJyBVEkiqIJA3NlLQsMkkVRJIqSEosFpJwU1yS7p5JYgWRpGIhScUiJSaCJAVCSqwgko/GAXeQlLDTahDKiiCVlRjgMR5GLisJ6e6eSag1glRrJKHWCMLdPUmKYUHKYUkK4pS0pWSSFnmSFNopMbUlqVRJwmeqTMIdQ0ksapKwSWaSJiJJ2I6X5AnGKM0wojjFGKXyZ5QmGVGcZYzCXUCnNCOJ4pRklNqnKLZPo7Dd75SmOlGcwYxSURXFOcwoTWJGaRYTpWnMZG3fMHTanMlOX/7p+b//enr+x7dfn59+O/Xn/cdvT/8+fTxFe4/xu9Rju7t2yxjT5q+nGm/T8smjv8SjcaG1H40z9vLRWEgNn3m0xLC5PqpXs0ftyPzRsavtR/Vq9qgdmT+a38XPp+9/iJPy4dVTFZdNPf/Pf/ub707lvP/47pu/nb4+xVuqM7lXjMuZnOxMXh/95e7RsuhRPcPnH9Uz/BxnLebR2+NbTAGX15v7F4/+cvdoLNrq7bxv8Kie4edPTszlO79cQ6+dsv+E2n/9xf5/ccq7eVjjWZ9+PX39/vpb8SRrvEaNGXmfScqwxgX7/tfTl3+p5/X8/uPp+3fT8tV57M/vnqOFLHWZ5uc5Hul/OL//6+nP709/f+xlatStuv8rqa+9TDSx/b/EUefyVXzf53d/5EX0vcQ8P8z7cDxeZP+A4PVVhnEd57cc/mvPPMSFHwvyqBb1OX6efv9rRH+JaFxjaogCvdTy6Qn66atzRNG7+eP8vIzzx9rX+N/041zmdX6a48TFNROrnfn5LS89130n8XOnbYif4lscrt9kHz+XoQ5D/OryY41frfFof5yK+pajmEq3TS+OYqrxve+XSTz3F/slucQBTfFuxAl4jmXZNo/7b9ZoLnGq4vfHP3AN6SDGPd4+cxAxBKY40eMSP49PccKH452vcSDrW15viOAcXl5XdZgupzRGyBbf5dNbXiMCZf8rHZ+M8Din83wddnEC46vn5cc3vExM2vt0+JlvpeTVEj/s15dTOMTVMw7L8esxvrqZ45oa3vJm7hP0VF58458OnnhP94ur7idiuJ6Qt7ynJea66TMXcqz351I/jB/2142Luezjen+o3y+viMHLRf1hH83zc/1wu8jjkEIu1wv9LQcWk9rymXhc4t3/uAzx+vvFfXd4P0Xo7Lux19+MN/NNZ2bc70G/vET2C+B42/fBHtPN5Rsv+893r/b/pcmbUwplbmRzdHJlYW0KZW5kb2JqCjEyIDAgb2JqCihTd2lzc1FSQmlsbCkKZW5kb2JqCjEzIDAgb2JqCihTd2lzc1FSQmlsbCkKZW5kb2JqCjE0IDAgb2JqCihEOjIwMjMwMzI5MDg1NzE4WikKZW5kb2JqCjE1IDAgb2JqCihTd2lzc1FSQmlsbCkKZW5kb2JqCjExIDAgb2JqCjw8Ci9Qcm9kdWNlciAxMiAwIFIKL0NyZWF0b3IgMTMgMCBSCi9DcmVhdGlvbkRhdGUgMTQgMCBSCi9BdXRob3IgMTUgMCBSCj4+CmVuZG9iagoxMCAwIG9iago8PAovVHlwZSAvRm9udAovQmFzZUZvbnQgL0hlbHZldGljYQovU3VidHlwZSAvVHlwZTEKL0VuY29kaW5nIC9XaW5BbnNpRW5jb2RpbmcKPj4KZW5kb2JqCjkgMCBvYmoKPDwKL1R5cGUgL0ZvbnQKL0Jhc2VGb250IC9IZWx2ZXRpY2EtQm9sZAovU3VidHlwZSAvVHlwZTEKL0VuY29kaW5nIC9XaW5BbnNpRW5jb2RpbmcKPj4KZW5kb2JqCjQgMCBvYmoKPDwKPj4KZW5kb2JqCjMgMCBvYmoKPDwKL1R5cGUgL0NhdGFsb2cKL1BhZ2VzIDEgMCBSCi9OYW1lcyAyIDAgUgo+PgplbmRvYmoKMSAwIG9iago8PAovVHlwZSAvUGFnZXMKL0NvdW50IDEKL0tpZHMgWzcgMCBSXQo+PgplbmRvYmoKMiAwIG9iago8PAovRGVzdHMgPDwKICAvTmFtZXMgWwpdCj4+Cj4+CmVuZG9iagp4cmVmCjAgMTYKMDAwMDAwMDAwMCA2NTUzNSBmIAowMDAwMDA5Mzg0IDAwMDAwIG4gCjAwMDAwMDk0NDEgMDAwMDAgbiAKMDAwMDAwOTMyMiAwMDAwMCBuIAowMDAwMDA5MzAxIDAwMDAwIG4gCjAwMDAwMDAzMDIgMDAwMDAgbiAKMDAwMDAwMDE3NCAwMDAwMCBuIAowMDAwMDAwMDU5IDAwMDAwIG4gCjAwMDAwMDAwMTUgMDAwMDAgbiAKMDAwMDAwOTE5OSAwMDAwMCBuIAowMDAwMDA5MTAxIDAwMDAwIG4gCjAwMDAwMDkwMTAgMDAwMDAgbiAKMDAwMDAwODg4NCAwMDAwMCBuIAowMDAwMDA4OTE0IDAwMDAwIG4gCjAwMDAwMDg5NDQgMDAwMDAgbiAKMDAwMDAwODk4MCAwMDAwMCBuIAp0cmFpbGVyCjw8Ci9TaXplIDE2Ci9Sb290IDMgMCBSCi9JbmZvIDExIDAgUgovSUQgWzwwZDY3ZmI3YmFmMzM0NmNhNDQwZTQ3NzZlMDU0OTJmNj4gPDBkNjdmYjdiYWYzMzQ2Y2E0NDBlNDc3NmUwNTQ5MmY2Pl0KPj4Kc3RhcnR4cmVmCjk0ODgKJSVFT0YK'", responseNode.get(Executor.ERRORS).get(0).asText());
		}
		finally
		{
			if (Objects.nonNull(is))
				is.close();
		}
	}

	@Test
	public void testDocumentInfoWithInvalidContentParameter() throws IOException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put("content", sourceFilename);

		ContentResponse response = client.POST("http://localhost:4567/fsl/Pdf.getDocumentInfo").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		System.out.println(responseNode.get(Executor.ERRORS).get(0).asText());
		assertEquals("Illegal base64 character 2e", responseNode.get(Executor.ERRORS).get(0).asText());
	}

	@Test
	public void testDocumentInfoWithFile() throws IOException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put("file", sourceFilename);

		ContentResponse response = client.POST("http://localhost:4567/fsl/Pdf.getDocumentInfo").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.OK, responseNode.get(Executor.STATUS).asText());
		JsonNode result = mapper.readTree(responseNode.get(Executor.RESULT).asText());
		assertEquals("SwissQRBill", result.get("author").asText());
		assertEquals("SwissQRBill", result.get("creator").asText());
		assertEquals("", result.get("keywords").asText());
		assertEquals("SwissQRBill", result.get("producer").asText());
		assertEquals("", result.get("subject").asText());
		assertEquals("", result.get("title").asText());
		assertEquals("2023/03/29 10:57:18", result.get("creationDate").asText());
		assertEquals("", result.get("modificationDate").asText());
	}

	@Test
	public void testDocumentInfoWithNotExistingFile() throws IOException, InterruptedException, TimeoutException, ExecutionException
	{
		ObjectNode requestNode = mapper.createObjectNode();
		requestNode.put("file", "gigi");

		ContentResponse response = client.POST("http://localhost:4567/fsl/Pdf.getDocumentInfo").
				header("Content-Type", "application/json").
				content(new StringContentProvider(requestNode.toString())).
				accept("application/json").
				send();
		
		JsonNode responseNode = mapper.readTree(response.getContentAsString());
		assertEquals(Executor.ERROR, responseNode.get(Executor.STATUS).asText());
		assertEquals(1, responseNode.get(Executor.ERRORS).size());
		assertEquals("not a file 'gigi'", responseNode.get(Executor.ERRORS).get(0).asText());
	}
}

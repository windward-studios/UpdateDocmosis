import com.docmosis.template.population.CompoundDataProvider;
import com.docmosis.template.population.DataProvider;
import com.docmosis.template.population.MemoryDataProvider;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import sun.misc.BASE64Encoder;

import javax.xml.parsers.*;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;

/*
 * Write out a Docmosis DataProvider object to an XML file.
 */
public class PublishDataProvider {

	private static Document dom;

	/**
	 * Write out a Docmosis DataProvider object to an XML file.
	 * @param provider The provider to write out as XML.
	 * @param xmlData The stream to write the XML to.
	 * @throws Exception Generally the data is using a class we don't handle.
	 */
	public static void publish(DataProvider provider, OutputStream xmlData) throws Exception {

		// create the DOM
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		DocumentBuilder db = dbf.newDocumentBuilder();
		// create instance of DOM
		dom = db.newDocument();

		// create the root element
		Element rootElement = dom.createElement("data");

		processProvider(provider, rootElement);

		dom.appendChild(rootElement);

		// save it to the output stream
		Transformer tr = TransformerFactory.newInstance().newTransformer();
		tr.setOutputProperty(OutputKeys.INDENT, "yes");
		tr.setOutputProperty(OutputKeys.METHOD, "xml");
		tr.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
		tr.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");

		// send DOM to file
		tr.transform(new DOMSource(dom), new StreamResult(xmlData));
		xmlData.flush();
	}

	/**
	 * Process a provider.
	 * @param provider Must be a MemoryDataProvider or CompoundDataProvider.
	 * @param rootElement The element in the DOM to write this provider's data to.
	 */
	private static void processProvider(DataProvider provider, Element rootElement) throws Exception {

		if (provider instanceof MemoryDataProvider) {
			writeProvider((MemoryDataProvider) provider, rootElement);
			return;
		}

		if (provider instanceof CompoundDataProvider) {
			DataProvider [] listProviders = getProviders((CompoundDataProvider)provider);
			for (DataProvider dataProvider : listProviders)
				processProvider(dataProvider, rootElement);
			return;
		}

		throw new UnsupportedOperationException("Do not know how to handle a " + provider.getClass().getName() + " DataProvider");
	}

	/**
	 * CompoundDataProvider has an array of providers but it's not a public field. So use reflectiont o get it.
	 * @param provider The CompoundDataProvider we pull the array of providers from.
	 * @return The DataProviders held by this CompoundDataProvider.
	 */
	private static DataProvider [] getProviders(CompoundDataProvider provider) throws IllegalAccessException {

		// we grab this by signature not name because obfuscitation can change the renaming between versions
		Field[] allFields = provider.getClass().getDeclaredFields();
		for (Field field : allFields)
			if (field.getType().isAssignableFrom(DataProvider[].class)) {
				field.setAccessible(true);
				return (DataProvider[]) field.get(provider);
			}
		throw new UnsupportedOperationException("Could not find list of providers array in CompoundDataProvider");
	}

	/**
	 * Write a provider to the XML file.
	 * @param provider The provider to write out.
	 * @param parentNode The element in the DOM to write this provider's data to.
	 */
	private static void writeProvider(MemoryDataProvider provider, Element parentNode) throws IOException {

		// for XML (and maybe others) it has the data in twice, as <key>dave</key> and as <key><value>dave</value></key>
		Map<String, String> mapKeys = new HashMap<>();

		// string values are written to this node
		Set props = provider.getStringKeys();
		if (props != null)
			for (Object item : props) {
				String key = item.toString();
				mapKeys.put(key, null);
				String value = provider.getString(key);

				Element element = dom.createElement(key);
				element.appendChild(dom.createTextNode(value));
				parentNode.appendChild(element);
			}

		// boolean values are written to this node
		props = provider.getBooleanKeys();
		if (props != null)
			for (Object item : props) {
				String key = item.toString();
				mapKeys.put(key, null);
				boolean value = provider.getBoolean(key);

				Element element = dom.createElement(key);
				element.appendChild(dom.createTextNode(Boolean.toString(value)));
				parentNode.appendChild(element);
			}

		// images written to this node - get binary so need to uuencode
		props = provider.getImageKeys();
		if (props != null)
			for (Object item : props) {
				String key = item.toString();
				mapKeys.put(key, null);
				InputStream image = provider.getImage(key);

				BASE64Encoder encoder = new BASE64Encoder();
				ByteArrayOutputStream out = new ByteArrayOutputStream();
				encoder.encode(image, out);
				String result = out.toString("UTF-8");

				Element element = dom.createElement(key);
				element.appendChild(dom.createTextNode(result));
				parentNode.appendChild(element);
			}

		// and now we dive in to child data
		props = provider.getDataProviderKeys();
		if (props != null)
			for (Object item : props) {
				String key = item.toString();
				if (mapKeys.containsKey(key))
					continue;
				int numProviders = provider.getDataProviderCount(key);
				for (int index = 0; index < numProviders; index++) {
					MemoryDataProvider childProvider = (MemoryDataProvider) provider.getDataProvider(key, index);
					Element element = dom.createElement(key);
					writeProvider(childProvider, element);
					parentNode.appendChild(element);
				}
			}
	}
}

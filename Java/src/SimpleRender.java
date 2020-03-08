import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import com.docmosis.SystemManager;
import com.docmosis.document.DocumentProcessor;
import com.docmosis.template.population.DataProviderBuilder;
import com.docmosis.template.population.MemoryDataProvider;
import com.docmosis.util.Configuration;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;

/**
 * A simple example showing how to convert Docmosis data to an XML file.
 */
public class SimpleRender
{
    public static void main(String[] args) throws Exception {

        // Use the DataProviderBuilder to build the data provider from a String array.
        DataProviderBuilder dpb = new DataProviderBuilder();
		
        dpb.add("date", "1776-07-04");
        dpb.add("message", "We hold these truths to be self-evident");
        File file = new File("../data/Repeating1Data.xml");
        dpb.addXMLFile(file.getCanonicalFile());

        file = new File("./data/plot.jpg");
        dpb.addImage("picture", new FileInputStream(file.getCanonicalFile()));

        FileOutputStream xmlData = new FileOutputStream("c:/test/data.xml");
        PublishDataProvider.publish(dpb.getDataProvider(), xmlData);
        xmlData.close();
    }
}
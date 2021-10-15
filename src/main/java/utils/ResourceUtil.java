package utils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.net.URISyntaxException;
import java.net.URL;

/**
 * Classe utilitáia para ler e escrever arquivos da pasta /resources.  <br>
 * Como utilizar: <br>
 * 1) Coloque o arquivo a ser lido em qualquer pasta do projeto, se quiser pode criar uma pasta /resource <br>
 * 2) Clique NO ARQUIVO e selecione "Build Path -> Use as Source Folder". <br>
 * Obs.: Quando se coloca o arquivo no Build Path ele ficará na raiz do JAR.
 * <br>
 * O método writeResource() escreve o arquivo dentro da pasta /resource pela IDE,          <br>
 * mas quando exporta em JAR os arquivos ficarão na raiz do JAR e serão apenas leitura     <br> 
 * o programa consegue ler mas não consegue escrever se estiver dentro do JAR.             <br>
 * normalmente o arquivo a compilação como dentro do JAR.                                  <br>
 * O m�todo readResource() retorna o conteúdo do arquivo dentro da pasta resource.         <br>
 * See <a href="https://mkyong.com/java/java-read-a-file-from-resources-folder/"> teste    </a>
 * 
 * @author Rodrigo Eggea
 * 
 * 
 */
public class ResourceUtil {

	/**
	 * Retorna um arquivo da pasta root do JAR.
	 * @return
	 * @throws FileNotFoundException 
	 * @throws IOException 
	 */
	public static String readFileString(String filename) throws FileNotFoundException {
		StringBuilder sb = new StringBuilder();
	    InputStream in = ResourceUtil.class.getResourceAsStream("/" + filename);
	    if(in == null) {
	    	throw new FileNotFoundException("File not found: " + filename);
	    }
        int content;
        try {
			while ((content = in.read()) != -1) {
			    sb.append((char)content);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
        return sb.toString();
	}
	
	public static InputStream readFile(String filename) {
		InputStream in = ResourceUtil.class.getResourceAsStream("/" + filename);
		return in;
	}
	
	/**
	 * Escreve arquivo na pasta resources. <br> 
	 * S� funciona durante a compila��o. N�o funciona depois de empacotado em um JAR. <br>
	 * 
	 * @param filename - Filename
	 * @throws FileNotFoundException
	 * @throws IOException
	 * @throws URISyntaxException
	 */
	public static void writeFile(String filename, String conteudo) {
		try { 
			URL url = ResourceUtil.class.getResource("/" + filename);
	    	File f = new File("resources/" + filename);
	    	PrintWriter pw = new PrintWriter(f);
	    	pw.write(conteudo);
	    	pw.close();
		} catch (Exception e) {
			// Vai dar erro se o programa estiver empacotado
			// dentro de um JAR e tentar escrever o resource.
			// Ignora o erro neste caso.
		}
	}
	
}

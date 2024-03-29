package exemplo;

import java.io.InputStream;

import leitorexcel.LeitorExcel;
import leitorexcel.LeitorExcel.EXTENSAO;
import utils.ResourceUtil;

/**
 * Exemplos de utilização da classe auxiliar LeitorExcel.
 * 
 * @author Rodrigo Eggea
 *
 */
public class MainResource {

	public static void main(String[] args) throws Exception {
		System.out.println("INICIO DO PROGRAMA");		
		System.out.println("LENDO ARQUIVO EXCEL: RDH_09OUT2021.xlsx");
		
		InputStream is = ResourceUtil.readFile("RDH_09OUT2021.xlsx");
		System.out.println("ARQUIVO CARREGADO!");
		
		System.out.println("LENDO ARQUIVO COM APACHE POI");
		LeitorExcel leitor = new LeitorExcel(is, EXTENSAO.XLSX);  // DENTRO DO JAR A LEITURA É LENTA.
		leitor.abrir(0);
		
		// Le uma string
		System.out.println("Lendo celula B8");
		System.out.println("   valor=" + leitor.getCelulaString("B8"));
		System.out.println();
		
		System.out.println("Lendo celula 3:8");
		System.out.println("   valor=" + leitor.getCelulaString(3, 8));
		System.out.println();
		
		// Le um valor usando coluna e linha
		System.out.println("Lendo a célula coluna=6 linha=56");
		System.out.println("   valor= " + leitor.getCelulaDouble(6,56));
		System.out.println();
		
		// Le um valor usando letra e numero
		System.out.println("Lendo a célula F56");
		System.out.println("   valor= " + leitor.getCelulaDouble("F56"));
		System.out.println();
		
		// Le uma coluna de valores
		System.out.println("Lendo uma coluna F48:F58");
		leitor.getColunaDouble("F48:F58").forEach(s -> System.out.println("    " + s));
		System.out.println();
		
		// Le uma coluna de valores
		System.out.println("Lendo a coluna 6:48-6:58");
		leitor.getColunaDouble(6,48,58).forEach(s -> System.out.println("    " + s));
		System.out.println();
		
		// Le uma linha de valores
		System.out.println("Lendo a linha F48:M48");
		leitor.getLinhaDouble("F48:M48").forEach(s -> System.out.print("    " + s));
		System.out.println();
		System.out.println();
		
		// Le uma linha de valores		
		System.out.println("Lendo a linha 6:48-13:48");
		leitor.getLinhaDouble(6,13,48).forEach(s -> System.out.print("    " + s));
		System.out.println();
		System.out.println();
		
		// Le uma matriz de valores
		{
			System.out.println("Lendo matriz de valores 6:48-7:50");
			Double[][] valores = leitor.getRangeDouble(6, 48, 7, 50);
			for(Double[] linha : valores) {
				for(Double celula : linha) {
					System.out.format("%10s", celula);
				}
				System.out.println();
			}
		}
		System.out.println();
		
		// Le uma matriz de valores
		{ 
			System.out.println("Lendo matriz de valores F48:G51");
			Double[][] valores = leitor.getRangeDouble("F48:H50");
			for(Double[] linha : valores) {
				for(Double celula : linha) {
					System.out.format("%10s", celula);
				}
				System.out.println();
			}
		}
		
		// Lendo outra planilha
		System.out.println("Abrindo planilha 1");
		leitor.abrir(1);
		
		// Lendo Dados
		System.out.println("Lendo titulo da planilha");
		System.out.println("   " + leitor.getCelulaString("C1"));
		
		System.out.println();
		System.out.println("FIM.");
	}
}

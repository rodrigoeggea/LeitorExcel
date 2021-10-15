package exemplo;

import java.nio.file.Path;
import java.nio.file.Paths;

import leitor.LeitorRdh;

public class MainLerArquivoRDH {

	public static void main(String[] args) throws Exception {
		System.out.println("INICIO DO PROGRAMA");
		
//		InputStream is = ResourceUtil.readFile("RDH_09OUT2021.xlsx");
//		LeitorRdh rdhExcel = new LeitorRdh(is, EXTENSAO.XLSX);
		Path pathRdh = Paths.get("c:/temp/RDH_09OUT2021.xlsx");
		LeitorRdh rdhExcel = new LeitorRdh(pathRdh);
		
		System.out.println("Vazao do posto 2=" + rdhExcel.getVazao(2));
		System.out.println(rdhExcel.getPostoVazao(2));
		System.out.println(rdhExcel.getPostoVazao(100));
		
	}
}

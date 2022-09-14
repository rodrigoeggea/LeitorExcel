package com.testes;
import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.InputStream;
import java.util.Arrays;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import leitorexcel.LeitorExcel;
import leitorexcel.LeitorExcel.EXTENSAO;
import utils.ResourceUtil;

class JunitTestaLeitorXls {
	private static InputStream is;
	private static LeitorExcel leitor;
	
	@BeforeAll
	static void setUpBeforeClass() throws Exception {
		is = ResourceUtil.readFile("RDH_09OUT2021.xlsx");
		leitor = new LeitorExcel(is, EXTENSAO.XLSX);
		leitor.abrir(0);
	}
	
	@Test
	void test1() throws Exception {
		assertEquals(34.0, leitor.getCelulaDouble("F8"));
		assertEquals(34.0, leitor.getCelulaDouble(6,8));
		
		assertEquals(0.27, leitor.getCelulaDouble("Z8"));
		assertEquals(0.27, leitor.getCelulaDouble(26,8));
		
		assertEquals(561.62, leitor.getCelulaDouble("O59"));
		assertEquals(561.62, leitor.getCelulaDouble(15,59));
	}
	
	@Test
	void test2() throws Exception {
		Double[] valoresTeste = {161.0, 164.0, 237.0, 238.0};
		assertEquals(Arrays.asList(valoresTeste), leitor.getColunaDouble(5, 48, 51));
		assertEquals(Arrays.asList(valoresTeste), leitor.getColunaDouble("E48:E51"));
	}
	
	@Test
	void test3() throws Exception {
		Double[][] valoresTeste = new Double[][]{
			{ 42.0, 59.0,  66.0},
			{ 28.0, 57.0,  45.0},
			{133.0, 52.0, 151.0},
			{155.0, 54.0, 169.0}};

			Double[][] result = leitor.getRangeDouble("F48:H51");
		
		for(int i=0; i<4; i++) {
			for(int j=0; j<3; j++) {
				assertEquals(valoresTeste[i][j], result[i][j]);
			}
			System.out.println();
		}
	}
	
	@Test
	void test4() throws Exception {
		Double[][] valoresTeste = new Double[][]{
			{ 42.0, 59.0,  66.0},
			{ 28.0, 57.0,  45.0},
			{133.0, 52.0, 151.0},
			{155.0, 54.0, 169.0}};

			Double[][] result = leitor.getRangeDouble(6,48, 8, 51);
		
		for(int i=0; i<4; i++) {
			for(int j=0; j<3; j++) {
				assertEquals(valoresTeste[i][j], result[i][j]);
			}
			System.out.println();
		}
	}
}

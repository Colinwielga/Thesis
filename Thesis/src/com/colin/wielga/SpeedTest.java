package com.colin.wielga;

import java.util.Date;
import java.util.Random;

public class SpeedTest {
	public static final int CMP = 0;
	public static final int QUICKCMP = 1;
	public static final int FASTCMP = 2;

	public static String rndString(char[] letter, int len) {
		Random r = new Random();
		String s = new String();
		for (int i = 0; i < len; i++) {
			s = s + letter[(int) (r.nextFloat() * letter.length)];
		}
		return s;
	}

	public static long test(int fnc, int len, char[] letter, int num) {
		long startTime = new Date().getTime();
		String a,b;
		if (fnc == CMP) {
			for (int i = 0; i < num; i++) {
				a = rndString(letter, len);
				b = rndString(letter, len);
				//System.out.print(a+" "+b +" | ");
				Cmp.cmp(a, b);
			}
		}
		if (fnc == QUICKCMP) {
			for (int i = 0; i < num; i++) {
				a = rndString(letter, len);
				b = rndString(letter, len);
				//System.out.print(a+" "+b +" | ");
				Cmp.qc_rapper(a, b);
			}
		}
		if (fnc == FASTCMP) {
			for (int i = 0; i < num; i++) {
				a = rndString(letter, len);
				b = rndString(letter, len);
				//System.out.print(a+" "+b +" | ");
				Cmp.fastCmp(a, b);
			}
		}
		long endTime = new Date().getTime();
		return endTime - startTime;
	}

}

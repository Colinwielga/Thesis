package com.colin.wielga;

import java.util.ArrayList;

public class CmpResult {

	public double score;
	public int[] raw;
	double unmatched;
	
	public CmpResult(double resultsum, int[] raw, int unmatched){
		this.score = resultsum;
		this.raw = raw;
		this.unmatched = unmatched;
	}

}

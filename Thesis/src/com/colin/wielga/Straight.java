package com.colin.wielga;

public class Straight {
	public int[] ystart;
	public int[] xstart;
	public int[] yend;
	public int[] xend;
	public int[] len;
	public int value;

	public Straight(int[] ystart, int[] xstart, int[] yend, int[] xend,
			int[] len) {
		this.ystart = ystart;
		this.xstart = xstart;
		this.yend = yend;
		this.xend = xend;
		this.len = len;
		value  =0;
		for (int i=0;i<this.len.length;i++){
			value+= Cmp.scale(this.len[i]);
		}
	}

	public Straight(int ystart, int xstart, int yend, int xend, int len) {
		this.ystart = new int[1];
		this.ystart[0] = ystart;
		this.xstart = new int[1];
		this.xstart[0] = xstart;
		this.yend = new int[1];
		this.yend[0] = yend;
		this.xend = new int[1];
		this.xend[0] = xend;
		this.len = new int[1];
		this.len[0] = len;
		value  =0;
		for (int i=0;i<this.len.length;i++){
			value+= Cmp.scale(this.len[i]);
		}
	}

	public Straight(Straight s1, Straight s2) {
		// maybe show a message if these can't coexists
		this.xstart = addAll(s1.xstart, s2.xstart);
		this.ystart = addAll(s1.ystart, s2.ystart);
		this.xend = addAll(s1.xend, s2.xend);
		this.yend = addAll(s1.yend, s2.yend);
		this.len = addAll(s1.len, s2.len);
		value  =0;
		for (int i=0;i<this.len.length;i++){
			value+= Cmp.scale(this.len[i]);
		}
	}

	public Straight() {
		ystart = new int[0];
		xstart = new int[0];
		yend = new int[0];
		xend = new int[0];
		len = new int[0];
		value=0;
	}

	private int[] addAll(int[] a, int[] b) {
		int[] result= new int[a.length+b.length];
		for (int i=0;i<a.length;i++){
			result[i] = a[i];
		}
		for (int i=0;i<b.length;i++){
			result[i+ a.length] = b[i];
		}
		return result;
	}

	public boolean coexist(Straight s) {
		// check for overlap
		for (int i = 0; i < s.len.length; i++) {
			for (int j = 0; j < len.length; j++) {
				if (s.xstart[i] <= xstart[j] && xstart[j] < s.xend[i]) {
					return false;
				} if (s.xstart[i] < xend[j] && xend[j] <= s.xend[i]) {
					return false;
				} if (s.ystart[i] <= ystart[j] && ystart[j] < s.yend[i]) {
					return false;
				} if (s.ystart[i] < yend[j] && yend[j] <= s.yend[i]) {
					return false;
				}
				if (xstart[j] <= s.xstart[i] && s.xstart[i] < xend[j]) {
					return false;
				} if (xstart[j] < s.xend[i] && s.xend[i] <= xend[j]) {
					return false;
				} if (ystart[j] <= s.ystart[i] && s.ystart[i] < yend[j]) {
					return false;
				} if (ystart[j] < s.yend[i] && s.yend[i] <= yend[j]) {
					return false;
				}
			}
		}
		// check to see if every sequence in one could be attached to a sequence in another 
		for (int i = 0; i < s.len.length; i++) {
			boolean ret = false;
			for (int j = 0; !ret && j < len.length; j++) {
				if ((s.xstart[i] == xend[j] && s.ystart[i] == yend[j]) || (s.xend[i] == xstart[j] && s.yend[i] == ystart[j])){
					ret =true;
				}				
			}
			if (ret){
				return false;
			}
		}
		return true;
	}

	public int value(){
		int result =0;
		for (int i=0;i<len.length;i++){
			result+= len[i]*2 - 1;
		}
		return result;
	}

	public Straight[] contains() {
		Straight[] ret = new Straight[xstart.length];
		for (int i=0;i<xstart.length;i++){
			ret[i] = new Straight(ystart[i],xstart[i],yend[i],xend[i], len[i]);
		}
		return ret;
	}

	public boolean holds(Straight target) {
		// TODO return true if this straight contains that straight
		// !coexists?
		return false;
	}
	
	public void print(){
		System.out.print("ystart ");
		for (int i=0;i<ystart.length;i++){
			System.out.print(" "+ystart[i]);
		}
		System.out.println();
		System.out.print("xstart ");
		for (int i=0;i<xstart.length;i++){
			System.out.print(" "+xstart[i]);
		}
		System.out.println();
		System.out.print("yend ");
		for (int i=0;i<yend.length;i++){
			System.out.print(" "+yend[i]);
		}
		System.out.println();
		System.out.print("xend ");
		for (int i=0;i<xend.length;i++){
			System.out.print(" "+xend[i]);
		}
		System.out.println();
	}
	
	public boolean eqls(Straight s){
		boolean in;
		if (s.xstart.length == this.xstart.length){
			for (int i=0; i<s.xstart.length;i++){
				in = false;
				for (int j=0 ; j<this.xstart.length;j++){
					if (s.xstart[i] == this.xstart[j]){
						in = true;
					}
				}
				if (!in){
					return false;
				}
			}
			for (int i=0; i<s.ystart.length;i++){
				in = false;
				for (int j=0 ; j<this.ystart.length;j++){
					if (s.ystart[i] == this.ystart[j]){
						in = true;
					}
				}
				if (!in){
					return false;
				}
			}
			for (int i=0; i<s.xend.length;i++){
				in = false;
				for (int j=0 ; j<this.xend.length;j++){
					if (s.xend[i] == this.xend[j]){
						in = true;
					}
				}
				if (!in){
					return false;
				}
			}
			for (int i=0; i<s.yend.length;i++){
				in = false;
				for (int j=0 ; j<this.yend.length;j++){
					if (s.yend[i] == this.yend[j]){
						in = true;
					}
				}
				if (!in){
					return false;
				}
			}
		}
		return true;
	}
}
package com.colin.wielga;

import java.util.ArrayList;

public class StraightCoexistThing {

	private static final int chunk = 100;
	private static final int GROWBY = 1000;
	private static final int STARTMAX = 1000;

	public ArrayList<Straight> straights = new ArrayList<Straight>();
	public int doneUpTo=0;
	ArrayList[] ray= new ArrayList[STARTMAX];
	private int[] at= new int[STARTMAX];
	StraightCoexistArray check = new StraightCoexistArray();
	
	public StraightCoexistThing(){
		for(int i =0; i<ray.length;i++){
			ray[i] = new ArrayList<int[]>();
		}
		for(int i =0; i<at.length;i++){
			at[i] = 0;
		}
		doneUpTo =0;	
	}

	public void reset() {
		for(int i =0; i<ray.length;i++){
			ray[i] = new ArrayList<int[]>();
		}
		for(int i =0; i<at.length;i++){
			at[i] = 0;
		}
		straights = new ArrayList<Straight>();
		doneUpTo =0;
		
	}

	public int size() {
		return straights.size();
	}

	public Straight getStraight(int l) {
		return straights.get(l);
	}

	public boolean get(int x, int y) {
		if (y > x){
			int c = x;
			x = y;
			y=c;
		}
		boolean chck = check.get(x,y);
		int upper = (ray[x].size()-1)*chunk + at[x];
		int lower = 0;
		int lookat;
		//System.out.print(upper + " ");
		while (upper > lower){
			lookat = (upper + lower)/2;
			if (((int[])ray[x].get(lookat/chunk))[lookat%chunk]==y){
				if (!chck){
					System.out.println(x+" "+y);
				}
				return true;
			}else if (lookat < 0 || (lookat/chunk == ray[x].size()-1 && lookat%chunk > at[x])){
				if (chck){
					System.out.println(x+" "+y);
				}
				return false;
			}else if (((int[])(ray[x].get(lookat/chunk)))[lookat%chunk]<y){
				upper =lookat -1;
			}else if (((int[])(ray[x].get(lookat/chunk)))[lookat%chunk]>y){
				lower =lookat +1;
			}else{
				System.out.println("something bad happened");
			}
		}
		if (chck){
			System.out.println(x+" "+y);
		}
		return false;
	}

	public void set(int x, int y, boolean value) {
		//System.out.print('-');
		if (y < x){
			int c = x;
			x = y;
			y=c;
		}
		check.set(y, x, value);
		if (x>=ray.length){
			System.out.println("i added a thing");
			ArrayList<int[]>[] copy = ray;
			ray = new ArrayList[ray.length + GROWBY];
			for (int i=0;i<copy.length;i++){
				ray[i] = copy[i];
			}
			for (int i = copy.length;i<ray.length;i++){
				ray[i] = ray[i] = new ArrayList<int[]>();
			}
			int[] atCopy = at;
			at = new int[ray.length];
			for (int i=0;i<atCopy.length;i++){
				at[i] = atCopy[i];
			}
			for (int i = atCopy.length;i<at.length;i++){
				at[i] = 0;
			}
		}
		int upper = (ray[x].size()-1)*chunk + at[x];
		int lower = 0;
		int lookat = (upper + lower)/2;;
		while (upper > lower){
			lookat = (upper + lower)/2;
			if (((int[])ray[x].get(lookat/chunk))[lookat%chunk]==y){
				upper = lookat;
				lower = lookat;
			}else if (lookat < 0 || (lookat/chunk == ray[x].size()-1 && lookat%chunk > at[x])){
				upper = lookat;
				lower = lookat;
			}else if (((int[])(ray[x].get(lookat/chunk)))[lookat%chunk]<y){
				upper =lookat -1;
			}else if (((int[])(ray[x].get(lookat/chunk)))[lookat%chunk]>y){
				lower =lookat +1;
			}else{
				System.out.println("something bad happened");
			}
		}
		if (ray[x].size() == 0){
			// if we have not put anything in ray[x] yet 
			((ArrayList<int[]>)ray[x]).add(new int[chunk]);
			((int[])ray[x].get(0))[0]= y;
			at[x] = at[x] +1;
		}else if (((int[])(ray[x].get(lookat/chunk)))[lookat%chunk]==y){
			if (!value){
				//System.out.println(x +" "+ y + " " + upper + " " + (ray[x].size()-1)*chunk + at[x]);
				for (int i = upper;i<(chunk*(ray[x].size()-1)) + at[x]-1;i++){
					//System.out.println(ray.length +" "+ x + " " + upper + " " + i+ " " + ray[x].size() +" "+ at[x]+" "+(upper+i+1)/chunk);
					((int[])(ray[x].get((i)/chunk)))[(i)%chunk]=((int[])(ray[x].get((i+1)/chunk)))[(i+1)%chunk];
				}
				at[x]=at[x]-1;
				if (at[x] == -1){
					at[x] = chunk -1;
					ray[x].remove(ray[x].size()-1);
				}
			}
		}else{
			if (value){
				if (upper<0){upper = 0;}
				at[x]=at[x]+1;
				if (at[x] >= chunk){
					at[x] = 0;
					((ArrayList<int[]>)ray[x]).add(new int[chunk]);
				}
				for (int i = 0;i<chunk*(ray[x].size()-1) + at[x] - upper - 1;i++){
					((int[])ray[x].get((upper+i+1)/chunk))[(upper+i+1)%chunk]=((int[])(ray[x].get((upper+i)/chunk)))[(upper+i)%chunk];
				}
				((int[])ray[x].get((upper)/chunk))[(upper)%chunk] = y;

			}
			
		}
		if (get(x,y) != value){
			System.out.println(x+" "+y+" "+value);
		}else{
			System.out.println(x+" "+y+" "+value+" good");
		}
		
	}

	public void insert(Straight s) {
		straights.add(s);
		
	}

	public int value(int x) {
		return straights.get(x).value();
	}

}

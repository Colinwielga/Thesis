package com.colin.wielga;

public class Analysis {
	
	private static final int SHITISBAD = 0;

	public static double standardDeviation(double[] mat){
		double ave = average(mat);
		double sum =0;
		for (int i=0;i<mat.length;i++){
			sum = sum + (mat[i] - ave)*(mat[i] - ave);
		}
		sum = Math.sqrt(sum/mat.length);
		return sum;
	}

	public static double average(double[] mat){
		double ret=0;
		for (int i=0;i<mat.length;i++){
			ret= ret + mat[i];
		}
		ret = ret/mat.length;
		return ret;
	}
	
	public static String getAdressEnd(String adress){
		String[] broken = adress.split("//");
		if (broken.length>1){
			return  broken[broken.length-2]  + "//" + broken[broken.length-1];
		}
		System.out.print(" SHIT IS BAD - get Adress End");
		return "this should not happen";
	}
	
	public static int getOrigPos(int chearerPos){
		String lookingfor = getAdressEnd(Runner.namePlag.get(chearerPos));
		for (int i=0;i<Runner.nameOrig.size();i++){
			if (getAdressEnd(Runner.nameOrig.get(i)).equals(lookingfor)){
				return i; 
			}
		}
		System.out.print(" SHIT IS BAD, could not find "+ lookingfor +" ");
		return SHITISBAD; 
	}
	
	public static void order(){
		// order sorts Runner.mat, Runner.nameOrig , Runner.result, Runner.cheaters, Runner.namePlag by the length of result and cheaters
		
		int max;
		int maxAt;
		String holder;
		Double holderDub;
		
		// lets start with the long side of Runner.mat , Runner.nameOrig, Runner.result
		for (int i=0;i<Runner.result.size();i++){
			max = 0;
			maxAt =-1;
			for (int j=i;j<Runner.result.size();j++){
				if (Runner.result.get(j).length() >max){
					max = Runner.result.get(j).length();
					maxAt = j;
				}
			}
			// result
			holder = Runner.result.get(maxAt);
			Runner.result.set(maxAt, Runner.result.get(i));
			Runner.result.set(i,holder);
			// nameOrig
			holder = Runner.nameOrig.get(maxAt);
			Runner.nameOrig.set(maxAt, Runner.nameOrig.get(i));
			Runner.nameOrig.set(i,holder);
			// mat
			for (int j = 0; j<Runner.mat.length;j++){
				holderDub = Runner.mat[j][maxAt];
				Runner.mat[j][maxAt] = Runner.mat[j][i];
				Runner.mat[j][i] = holderDub;
			}
		}
		
		
		// now the short side side of Runner.mat , Runner.namePlag, Runner.cheaters
		for (int i=0;i<Runner.cheaters.size();i++){
			max = 0;
			maxAt =-1;
			for (int j=i;j<Runner.cheaters.size();j++){
				if (Runner.cheaters.get(j).length() >max){
					max = Runner.cheaters.get(j).length();
					maxAt = j;
				}
			}
			// result
			holder = Runner.cheaters.get(maxAt);
			Runner.cheaters.set(maxAt, Runner.cheaters.get(i));
			Runner.cheaters.set(i,holder);
			// nameOrig
			holder = Runner.namePlag.get(maxAt);
			Runner.namePlag.set(maxAt, Runner.namePlag.get(i));
			Runner.namePlag.set(i,holder);
			// mat
			for (int j = 0; j<Runner.mat[0].length;j++){
				holderDub = Runner.mat[maxAt][j];
				Runner.mat[maxAt][j] = Runner.mat[i][j];
				Runner.mat[i][j] = holderDub;
			}
		}
		//cool we are done
	}
	
	
	//i was thinking about using
	int partition(int arr[], int left, int right){
	      int i = left, j = right;
	      int tmp;
	      int pivot = arr[(left + right) / 2];
	     
	      while (i <= j) {
	            while (arr[i] < pivot)
	                  i++;
	            while (arr[j] > pivot)
	                  j--;
	            if (i <= j) {
	                  tmp = arr[i];
	                  arr[i] = arr[j];
	                  arr[j] = tmp;
	                  i++;
	                  j--;
	            }
	      };
	     
	      return i;
	}
	 
	void quickSort(int arr[], int left, int right) {
	      int index = partition(arr, left, right);
	      if (left < index - 1)
	            quickSort(arr, left, index - 1);
	      if (index < right)
	            quickSort(arr, index, right);
	}
	
	public static int numberFromWinner(double[] mat , int pos){
		int ret =1;
		for (int i = 0; i<mat.length;i++){
			if (i != pos && mat[pos] <= mat[i]){
				ret++;
			}
		}
		return ret;
	}
	
	public static double percentile(double[] in , int pos){
		return 100-((numberFromWinner(in,pos)*100)/in.length);
	}
	
	
	
	public static void printAll(){
		for (int i=0;i<Runner.mat.length;i++){
			//print what line we are looking at
			System.out.print(i);
			//print the orig
			System.out.print(","+Runner.mat[i][getOrigPos(i)]);
			//print the average
			System.out.print(","+average(Runner.mat[i]));
			//Print the std
			System.out.print(","+standardDeviation(Runner.mat[i]));
			//print the percentile
			System.out.print(","+numberFromWinner(Runner.mat[i],getOrigPos(i)));
		
			System.out.println();
		}
		
	}
	
	
}

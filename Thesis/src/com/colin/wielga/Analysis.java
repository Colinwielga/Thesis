package com.colin.wielga;

import java.util.ArrayList;

public class Analysis {

	private static final int SHITISBAD = 0;
	private static final int LINEAR = 1;
	private static final int FOURTH = 4;
	private static final int THRID = 3;
	private static final int QUAD = 2;
	private static final int EXP = 5;
	private static final int FAC = 6;
	private static final int NONE = 0;

	public static double standardDeviation(double[] mat) {
		double ave = average(mat);
		double sum = 0;
		for (int i = 0; i < mat.length; i++) {
			sum = sum + (mat[i] - ave) * (mat[i] - ave);
		}
		sum = Math.sqrt(sum / mat.length);
		return sum;
	}

	public static double average(double[] mat) {
		double ret = 0;
		for (int i = 0; i < mat.length; i++) {
			ret = ret + mat[i];
		}
		ret = ret / mat.length;
		return ret;
	}

	public static String getAdressEnd(String adress) {
		String[] broken = adress.split("//");
		if (broken.length > 1) {
			return broken[broken.length - 2] + "//" + broken[broken.length - 1];
		}
		System.out.print(" SHIT IS BAD - get Adress End");
		return "this should not happen";
	}

	public static int getOrigPos(int chearerPos) {
		String lookingfor = getAdressEnd(Runner.namePlag.get(chearerPos));
		for (int i = 0; i < Runner.nameOrig.size(); i++) {
			if (getAdressEnd(Runner.nameOrig.get(i)).equals(lookingfor)) {
				return i;
			}
		}
		System.out.print(" SHIT IS BAD, could not find " + lookingfor + " ");
		return SHITISBAD;
	}

	public static void order() {
		// order sorts Runner.mat, Runner.nameOrig , Runner.result,
		// Runner.cheaters, Runner.namePlag by the length of result and cheaters

		int max;
		int maxAt;
		String holder;
		Double holderDub;

		// lets start with the long side of Runner.mat , Runner.nameOrig,
		// Runner.result
		for (int i = 0; i < Runner.result.size(); i++) {
			max = -1;
			maxAt = -1;
			for (int j = i; j < Runner.result.size(); j++) {
				if (Runner.result.get(j).length() > max) {
					max = Runner.result.get(j).length();
					maxAt = j;
				}
			}
			// result
			holder = Runner.result.get(maxAt);
			Runner.result.set(maxAt, Runner.result.get(i));
			Runner.result.set(i, holder);
			// nameOrig
			holder = Runner.nameOrig.get(maxAt);
			Runner.nameOrig.set(maxAt, Runner.nameOrig.get(i));
			Runner.nameOrig.set(i, holder);
			// mat
			for (int j = 0; j < Runner.mat.length; j++) {
				holderDub = Runner.mat[j][maxAt];
				Runner.mat[j][maxAt] = Runner.mat[j][i];
				Runner.mat[j][i] = holderDub;
			}
		}

		// now the short side side of Runner.mat , Runner.namePlag,
		// Runner.cheaters
		for (int i = 0; i < Runner.cheaters.size(); i++) {
			max = -1;
			maxAt = -1;
			for (int j = i; j < Runner.cheaters.size(); j++) {
				if (Runner.cheaters.get(j).length() > max) {
					max = Runner.cheaters.get(j).length();
					maxAt = j;
				}
			}
			// cheaters
			holder = Runner.cheaters.get(maxAt);
			Runner.cheaters.set(maxAt, Runner.cheaters.get(i));
			Runner.cheaters.set(i, holder);
			// namePlag
			holder = Runner.namePlag.get(maxAt);
			Runner.namePlag.set(maxAt, Runner.namePlag.get(i));
			Runner.namePlag.set(i, holder);
			// mat
			for (int j = 0; j < Runner.mat[0].length; j++) {
				holderDub = Runner.mat[maxAt][j];
				Runner.mat[maxAt][j] = Runner.mat[i][j];
				Runner.mat[i][j] = holderDub;
			}
		}
		// cool we are done
	}

	// i was thinking about using
	int partition(int arr[], int left, int right) {
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
		}
		;

		return i;
	}

	void quickSort(int arr[], int left, int right) {
		int index = partition(arr, left, right);
		if (left < index - 1)
			quickSort(arr, left, index - 1);
		if (index < right)
			quickSort(arr, index, right);
	}

	public static int numberFromWinner(double[] mat, int pos) {
		int ret = 1;
		for (int i = 0; i < mat.length; i++) {
			if (i != pos && mat[pos] <= mat[i]) {
				ret++;
			}
		}
		return ret;
	}

	public static double percentile(double[] in, int pos) {
		return 100 - ((numberFromWinner(in, pos) * 100) / in.length);
	}

	public static void finalscore() {
		double ave;
		for (double match = 0; match < 7; match+=.1) {
			System.out.print(match+",");
			for (double unmatch = 0; unmatch < 7; unmatch+=.1) {
				ave=0;
				for (int i = 0; i < Runner.rawScores.length; i++) {
					for (int j = 0; j < Runner.rawScores[i].length; j++) {
						Runner.mat[i][j] = score(Runner.rawScores[i][j],match,unmatch);
					}
					ave = ave + numberFromWinner(Runner.mat[i], getOrigPos(i));
				}
				ave = ave / Runner.rawScores.length;
				System.out.print(ave+",");
			}
			System.out.println();
		}
	}

	private static double score(CmpResult raw, double code, double un) {
		// TODO Auto-generated method stub
		double ret = 0;
		// add up matched
		
		for (int i = 0; i < raw.raw.length; i++) {
			ret = ret + (raw.raw[i]/Math.abs(raw.raw[i]))*Math.pow(Math.abs(raw.raw[i]), code);
		}
		ret = ret - (raw.unmatched/Math.abs(raw.unmatched))*Math.pow(Math.abs(raw.unmatched), un);
		
//		if (code == LINEAR) {
//			for (int i = 0; i < raw.raw.length; i++) {
//				result = result + raw.raw[i];
//			}
//		} else if (code == QUAD) {
//			for (int i = 0; i < raw.raw.length; i++) {
//				result = result + Math.pow(raw.raw[i], 2);
//			}
//		} else if (code == THRID) {
//			for (int i = 0; i < raw.raw.length; i++) {
//				result = result + Math.pow(raw.raw[i], 3);
//			}
//		} else if (code == FOURTH) {
//			for (int i = 0; i < raw.raw.length; i++) {
//				result = result + Math.pow(raw.raw[i], 4);
//			}
//		} else if (code == EXP) {
//			for (int i = 0; i < raw.raw.length; i++) {
//				result = result + Math.pow(2, raw.raw[i]);
//			}
//		} else if (code == FAC) {
//			for (int i = 0; i < raw.raw.length; i++) {
//				result = result + fac(raw.raw[i]);
//			}
//		} else if (code == NONE) {
//			// do nothing
//		}
		// add up unmatched
//		if (un == LINEAR) {
//			result = result - raw.unmatched;
//		} else if (un == QUAD) {
//			result = result - Math.pow(raw.unmatched, 2);
//		} else if (un == THRID) {
//			result = result - Math.pow(raw.unmatched, 3);
//		} else if (un == FOURTH) {
//			result = result - Math.pow(raw.unmatched, 4);
//		} else if (un == EXP) {
//			result = result - Math.pow(2, raw.unmatched);
//		} else if (un == FAC) {
//			result = result - fac((int) raw.unmatched);
//		} else if (un == NONE) {
//			// Do nothing
//		}

		return ret;
	}

	private static double fac(Integer in) {
		if (in <= 1) {
			return 1;
		}
		return fac(in - 1) * in;
	}

	public static void printAll() {
		for (int i = 0; i < Runner.mat.length; i++) {
			// print what line we are looking at
			System.out.print(getAdressEnd(Runner.namePlag.get(i)));
			// print the len of orig
			System.out.print("," + Runner.cheaters.get(i).length());
			// print the orig
			System.out.print("," + Runner.mat[i][getOrigPos(i)]);
			// print the score of the winner
			System.out.print("," + winningScore(i));
			// print the average
			System.out.print("," + average(Runner.mat[i]));
			// Print the std
			System.out.print("," + standardDeviation(Runner.mat[i]));
			// print the rank
			System.out.print(","
					+ numberFromWinner(Runner.mat[i], getOrigPos(i)));
			// rank number better than +1
			System.out.print(","
					+ (numberBetterThan(Runner.mat[i], getOrigPos(i)))+1);
			System.out.println();
		}

	}

	private static int numberBetterThan(double[] mat, int pos) {
		int ret = 0;
		for (int i = 0; i < mat.length; i++) {
			if (i != pos && mat[pos] < mat[i]) {
				ret++;
			}
		}
		return ret;
	}

	private static double winningScore(int i) {
		// TODO Auto-generated method stub
		double max = Runner.mat[i][0];
		for (int j =1; j< Runner.mat[i].length;j++){
			if (Runner.mat[i][j] >max){
				max = Runner.mat[i][j];
			}
		}
		return max;
	}

}

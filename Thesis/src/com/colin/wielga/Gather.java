package com.colin.wielga;

public class Gather implements Runnable {
	Thread t;
	String data;
	String name;
	boolean working = false;

	public Gather(String name, String data) {
		this.data = data;
		this.name = name;
		// Create a new thread
		t = new Thread(this, name);
		t.start(); // Start the thread
	}

	@Override
	public void run() {
		// TODO Auto-generated method stub
		working = true;

		// find the next open dude
		for (int i = 0; i < Runner.open.length; i++) {
			for (int j = 0; j < Runner.open[i].length; j++) {
				if (Runner.open[i][j] == Runner.WAITING) {
					Runner.open[i][j] = Runner.WORKING;
					// System.out.println(name + " is working " + i + "/"
					// + Runner.cheaters.size() + " cheaters: "
					// + Runner.cheaters.get(i));
					// System.out.println(name + " is working " + j + "/"
					// + Runner.result.size() + " result: "
					// + Runner.result.get(j));
					//System.out.println(i + "," + j);
					//Runner.open[i][j] = Runner.TOWRITE;
					//Runner.mat[i][j] = Cmp.countCmp(Runner.cheaters.get(i), Runner.result.get(j));
					//Runner.mat[i][j] = Alingment.globalAl(Runner.cheaters.get(i),Runner.result.get(j));
					
					CmpResult q3 = Cmp.qc3(Runner.cheaters.get(i),Runner.result.get(j));
					Runner.rawScores[i][j] = q3;
					Runner.mat[i][j] = q3.score;
					
					// int q1 = Cmp.qc_rapper(Runner.result.get(i),
					// Runner.cheaters.get(j));

					// System.out.println("qc3 " + q3 +" q1 "+ q1+" of "+ i
					// +" "+ j);

					Runner.open[i][j] = Runner.TOWRITE;
					// if (name.equals("scribe")) {
					// Runner.TryWrite(data);
					// }
				}
			}
			if (name.equals("postOp")) {
				System.out.println(i + " / " + Runner.open.length);
			}
		}
		// Runner.TryWrite(data);
		System.out.println(name + " is done");

		working = false;

		if (name.equals("postOp")) {
			boolean analysis= false;
			while (!analysis){
				analysis = true;
				for (int i = 0; analysis && i < Runner.gathers.size(); i++) {
					if (Runner.gathers.get(i).working) {
						analysis = false;
					}
				}
			}
			if (analysis) {
				System.out.println("writing");
				// Analysis.order();
				// Runner.writeAll(Runner.mat, data);
				System.out.println("Analysing");
				//Analysis.finalscore();
				Analysis.printAll();
			}
		}
	}

}
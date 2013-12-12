package test;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;

public class test {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		System.out.println((int)Math.ceil((double)242/50));
		
		/*String ss = "sdf";
		try{
			Integer.valueOf(ss);
		}catch(NumberFormatException e){
			System.out.println(ss);
		}*/
		
		Map<Integer , Integer[]> calcul = new LinkedHashMap<Integer , Integer[]>();
       
        final int num = 30;
		for(int j = 0 ; j<10 ; j++){
        	
        	Integer[] values = new Integer[num];
        	for (int i = 0; i < values.length; i++) {
				values[i] = 0;
			}
        	calcul.put(j, values);
        }
		calcul.get(1)[1] = 100;
		System.out.println(calcul.get(1)[1]);
		
		Map<Integer , Integer> test = new LinkedHashMap<Integer, Integer>();
		for(int j = 0 ; j<10 ; j++){
			test.put(j, j);
        }
		//test.get(1) = 100;
		System.out.println(test.get(1));
	}

}

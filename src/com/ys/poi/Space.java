package com.ys.poi;

public class Space {

    public static void main(String[] args) {
        // TODO Auto-generated method stub
        String str = "   \n雪   人   大  肚  子 一  挺";
        System.out.println(str.startsWith(" "));
        for(;str.startsWith(" ");){
            
        }
        System.out.println(str.replaceFirst("\\s*","替换"));
    }

}

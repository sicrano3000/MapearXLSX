package com.example.mapearxlsx;

import java.util.*;

public class GenerateRandomSequence {

    public static void main(String[] args) {
        var jogos = new HashSet<>();

        for (int i = 0; i < 30; i++) {
            jogos.add(generateRandomSequence(15));
        }

        jogos.forEach(System.out::println);
    }

    public static HashSet<Object> generateRandomSequence(int size) {
        // Cria um conjunto vazio para armazenar os números aleatórios
        var sequence = new HashSet<>();

        // Gera 15 números aleatórios entre 0 e 60
        while (sequence.size() < size) {
            sequence.add((int) (Math.random() * 25) + 1);
        }

        // Retorna a sequência aleatória
        return sequence;
    }

}

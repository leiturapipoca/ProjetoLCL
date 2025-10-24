# ProjetoLCL
O PROJETO DO QUARTUS II  É "projeto_final_cpa" e a bomba é o arquivo "testando_maquina_de_estados.bdf"


# 💣 Projeto Bomba Relógio em FPGA

[cite_start]Este projeto, desenvolvido para a disciplina de Laboratório de Circuitos Lógicos [cite: 3][cite_start], consiste na implementação de um jogo digital de "Bomba Relógio" [cite: 5] [cite_start]em uma placa de desenvolvimento FPGA Altera DE2[cite: 50]. [cite_start]O objetivo é criar um sistema interativo que desafia o jogador a descobrir uma senha de 6 bits antes que um cronômetro de três minutos chegue a zero[cite: 11].

## 🎯 O Desafio

O jogo simula o desarmamento de uma bomba. Um jogador define uma senha secreta e inicia o cronômetro. Outro jogador deve então adivinhar essa senha. A cada tentativa incorreta, o sistema fornece dicas inteligentes para ajudar na próxima tentativa, enquanto o tempo continua correndo. [cite_start]Se o tempo se esgotar, a bomba "explode"[cite: 13]. Se a senha for adivinhada a tempo, ela é desarmada com sucesso.

## ⚙️ Funcionalidades Implementadas

* [cite_start]**Definição da Senha:** Ao pressionar o botão "Start", uma senha de 6 bits é capturada das chaves `SW[17:12]` e armazenada internamente[cite: 28].
* [cite_start]**Contador Regressivo:** Um cronômetro de 3 minutos (iniciado em 2:59) é ativado e exibido nos displays de 7 segmentos da placa[cite: 11, 28].
* [cite_start]**Sistema de Tentativas:** O jogador insere uma senha de 6 bits com as chaves `SW[5:0]` e a submete pressionando o botão "ENTER"[cite: 29].
* [cite_start]**Dicas Inteligentes:** Em caso de erro, o sistema oferece duas dicas cruciais para o jogador[cite: 37]:
    * [cite_start]**Dica Lógica:** O resultado das operações bit-a-bit `AND` e `OR` entre a senha correta e a tentativa é mostrado nos displays `HEX7` e `HEX6`[cite: 38].
    * [cite_start]**Dica de Posição:** O número de bits corretos na tentativa é exibido em uma barra de progresso nos LEDs vermelhos (`LEDR[5:0]`)[cite: 39].
* [cite_start]**Condições de Vitória e Derrota:** O jogo termina quando a senha é acertada (congelando o tempo e acendendo os LEDs verdes [cite: 36][cite_start]) ou quando o tempo se esgota, ativando um efeito visual de "explosão" com os LEDs e displays[cite: 42].

## 🛠️ Aprendizados e Desafios: Do Software ao Silício

Implementar este projeto foi muito além de apenas criar um jogo; foi um exercício prático profundo sobre o funcionamento de sistemas digitais no nível mais fundamental.

### A Utilidade do Verilog como Descrição de Hardware

Diferente de linguagens de programação convencionais como C++ ou Java, Verilog não é uma linguagem de programação, e sim uma **Linguagem de Descrição de Hardware (HDL)**. Seu propósito não é dar uma sequência de comandos para um processador executar, mas sim **descrever circuitos digitais**.

Neste projeto, o Verilog foi essencial para descrever como os componentes (portas lógicas, flip-flops, contadores, comparadores) se conectam e se comportam. Em vez de escrever um loop `for`, descrevemos um contador que incrementa a cada pulso de clock. [cite_start]Essa descrição é então usada pelo software de síntese (Quartus-II [cite: 50]) para gerar uma configuração de hardware real que é gravada no chip FPGA.

### Programação em Baixo Nível e o Paradigma Paralelo

O maior desafio ao trabalhar com HDLs é a mudança de um pensamento sequencial para um **pensamento paralelo**. Em um processador, as instruções rodam uma após a outra. Em um FPGA, tudo acontece ao mesmo tempo. O contador de tempo, o comparador de senhas e o gerador de dicas funcionam **simultaneamente**, em sincronia com um sinal de clock global. A lógica do jogo foi implementada usando **máquinas de estados finitos (FSMs)** para controlar o fluxo (estado inicial, estado de contagem, estado de verificação, estado de vitória/derrota), garantindo que as transições ocorressem de forma ordenada.

### Entendimento Profundo da Interação com o Hardware

Este projeto força uma conexão direta e inseparável entre o código e o hardware físico. Cada funcionalidade exigiu um entendimento de como os periféricos da placa DE2 funcionam:
* **Entradas:** Foi necessário ler o estado dos botões (`KEY`) e das chaves (`SW`), lidando com questões como *debouncing* para garantir uma leitura limpa.
* **Saídas:** Controlar os `LEDs` é simples, mas para os **displays de 7 segmentos**, foi preciso criar um decodificador que converte um número binário de 4 bits no padrão de 7 bits necessário para acender os segmentos corretos e formar um dígito hexadecimal.

Isso solidifica o entendimento de que cada linha de HDL resulta em **circuitos reais**, que consomem recursos físicos no chip e operam com restrições de tempo do mundo real. É a tradução direta da lógica abstrata para o silício.

module maquina(
	input [4:0]  entry,
	output  reg [7:0]  saidas
	
);
	
always @ (*) // 000(1) 001(2) 010(3) 100(4) 
 begin //saida[2] - saida[1] - saida[0] - carro_verde - carro_amarelo - carro_vermelho - pedestre-verde - pedestre-vermelho
casex (entry)
	5'b 10xxx: saidas = 8'b 10100000;//amarelo piscando
	5'b 10101: saidas = 8'b 11000000;
	5'b 10110: saidas = 8'b 10101000;
	5'b 11xxx: saidas = 8'b 00010001; //estado inicial
	5'b 01xxx: saidas = 8'b 00010001;//estado inicial tambem
	5'b 00000: saidas = 8'b 00110001;//estado inicial : vai para 001///// ESTADO 1
	5'b 00001: saidas = 8'b 01001001;//amarelo/vermelho vai para 010 ///// ESTADO 2
	5'b 00010: saidas = 8'b 10000110;//vermelho/verde, vai para 100///// ESTADO 3 
	5'b 00100: saidas = 8'b 00000100;//vermelho/pisca vermelho ///// ESTADO 4 
	default : saidas = 8'b 00011111; //...
	endcase
	end
	endmodule

module password_comparator (
    input  logic        clk,          // Clock do sistema
    input  logic [5:0]  gabarito,     // Senha correta
    input  logic [5:0]  resposta,     // Tentativa atual
    input  logic        enter,        // Botão enter (1 = pressionado)
    output logic [5:0]  LEDR          // Saída LEDR[5:0] (DE2)
);

    logic [5:0] resultado_reg;       // Registro para manter o estado
    
    always_ff @(posedge clk) begin
        if (enter) begin  // Somente atualiza quando enter é pressionado
            resultado_reg <= '0;     // Zera todos os LEDs
            for (int i = 0; i < 6; i++) begin
                // Acende LEDs progressivamente para cada bit correto
                if ((i < 6) && (gabarito[i] == resposta[i])) begin
                    resultado_reg[i] <= 1'b1;
                end
            end
        end
    end

    assign LEDR = resultado_reg;  // Conecta diretamente aos LEDs

endmodule
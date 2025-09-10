module led_pisca (
    input  logic clk,        // clock lento (ex: 1Hz)
    input  logic reset,      // reset global
    input  logic fim,        // pisca se derrota
    input  logic win,        // acende fixo se vitória
    output logic [11:0] ledr // LEDR[17:6]
);

    logic estado_led;

    always_ff @(posedge clk or posedge reset) begin
        if (reset)
            estado_led <= 0;
        else if (fim)
            estado_led <= ~estado_led; // pisca
        else
            estado_led <= 0; // garante 0 se não está em fim
    end

    // Seleção final dos LEDs
    always_comb begin
        if (win)
            ledr = 12'b1111_1111_1111; // vitória → ligado fixo
        else if (fim)
            ledr = estado_led ? 12'b1111_1111_1111 : 12'b0000_0000_0000; // derrota → pisca
        else
            ledr = 12'b0000_0000_0000; // jogo em andamento → desligado
    end

endmodule

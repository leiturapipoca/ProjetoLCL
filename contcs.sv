module contcs (
    input  logic clk,          // pulso de 10ms vindo do contador de milissegundos
    output logic [3:0] q,      // saída do contador (0 a 9)
    output logic inc           // pulso de incremento para próximo estágio (ex: décimos)
);

always_ff @(posedge clk) begin
    if (q == 9) begin
        q   <= 0;
        inc <= 1;
    end else begin
        q   <= q + 1;
        inc <= 0;
    end
end

endmodule

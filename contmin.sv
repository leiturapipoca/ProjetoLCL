module contmin (
    input logic clk,
    input logic reset,
    input logic enable,           // recebe dec dos segundos
    output logic fim,
    output logic [3:0] ds,
    output logic [3:0] us
);

logic [5:0] q;

always_ff @(posedge clk or posedge reset) begin
    if (reset) begin
        q   <= 3;
        fim <= 0;
    end else if (enable) begin
        if (q == 0) begin
            q   <= 0;
            fim <= 1; // sinal de que acabou
        end else begin
            q   <= q - 1;
            fim <= 0;
        end
    end else begin
        fim <= 0;
    end
end

assign ds = q / 10;
assign us = q % 10;

endmodule

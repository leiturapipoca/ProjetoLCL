module contseg_min (
    input logic clk,
    input logic reset,
    input logic enable,             // usa inc como enable
    output logic dec,
    output logic [3:0] ds,
    output logic [3:0] us
);

logic [5:0] q;

always_ff @(posedge clk or posedge reset or posedge enable) begin
    if (reset) begin
        q   <= 59;
        dec <= 0;
    end else if (enable) begin
        if (q == 0) begin
            q   <= 59;
            dec <= 1;
        end else begin
            q   <= q - 1;
            dec <= 0;
        end
    end else begin
        dec <= 0; // importante para não gerar dec spurious
    end
end

assign ds = q / 10;
assign us = q % 10;

endmodule

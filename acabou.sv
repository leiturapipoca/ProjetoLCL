
module acabou (
    input  logic [1:0] entrada,  
    input  logic       clk,
    input  logic       reset,
    output logic       ganhou,
    output logic       perdeu
);

    always_ff @(posedge clk or posedge reset) begin
        if (reset) begin
            ganhou <= 0;
            perdeu <= 0;
        end else begin
            case (entrada)
                2'b00, 2'b11: begin
                    
                    ganhou <= ganhou;
                    perdeu <= perdeu;
                end
                2'b01: begin
                    
                    ganhou <= 1;
                    perdeu <= 0;
                end
                2'b10: begin
                    
                    ganhou <= 0;
                    perdeu <= 1;
                end
            endcase
        end
    end

endmodule

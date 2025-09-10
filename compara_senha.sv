module compara_senha (
    input  logic        clk,         
    input  logic        reset,       
    input  logic        en,          
    input  logic [5:0]  senha,       
    input  logic [5:0]  tentativa,   
    output logic [5:0]  leds         
);

    
    logic [2:0] corretos;
    always_comb begin
        corretos = 3'd0;
        for (int i = 0; i < 6; i++) begin
            if (senha[i] == tentativa[i])
                corretos++;
        end
    end

    
    always_ff @(posedge clk) begin
        if (reset) begin
            leds <= 6'b0;
        end
        else if (en) begin
            case (corretos)
                3'd0: leds <= 6'b000000;
                3'd1: leds <= 6'b000001;
                3'd2: leds <= 6'b000011;
                3'd3: leds <= 6'b000111;
                3'd4: leds <= 6'b001111;
                3'd5: leds <= 6'b011111;
                3'd6: leds <= 6'b111111;
                default: leds <= 6'b000000;
            endcase
        end
        
    end

endmodule

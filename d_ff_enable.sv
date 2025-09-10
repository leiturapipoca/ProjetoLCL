module d_ff_enable (
    input  logic clk,       
    input  logic reset,     
    input  logic en,        
    input  logic d,         
    output logic q          
);

    always_ff @(posedge clk ,posedge en, posedge reset) begin
        if (reset)
            q <= 1'b0;       
        else if (en)
            q <= d;          
        
    end

endmodule

module contms(
	input logic clk, rst,
	output logic inc,
	output logic [3:0] q
);

always_ff @(posedge clk, posedge rst)
	if (rst) begin
		q <= 0;
		inc <= 0;
	end else if (q == 9) begin
		q <= 0;
		inc <= 1;
	end else begin
		q <= q + 1;
		inc <= 0;
	end

endmodule

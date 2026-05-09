import xlsxwriter

def create_final_answer_key():
    workbook = xlsxwriter.Workbook('LastName_FinalExam_Key.xlsx')
    
    # Common formats
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center'})
    bold_fmt = workbook.add_format({'bold': True})
    num_fmt = workbook.add_format({'num_format': '0.000000', 'border': 1})
    pct_fmt = workbook.add_format({'num_format': '0.00%', 'border': 1})
    border_fmt = workbook.add_format({'border': 1})
    
    # --- P1_Integration ---
    ws1 = workbook.add_worksheet('P1_Integration')
    ws1.set_column('A:F', 15)
    
    ws1.write('A1', 'Numerical Integration (n=10)', bold_fmt)
    ws1.write('A2', 'a =')
    ws1.write('B2', 0)
    ws1.write('A3', 'b =')
    ws1.write('B3', 0.8)
    ws1.write('A4', 'n =')
    ws1.write('B4', 10)
    ws1.write('A5', 'h =')
    ws1.write_formula('B5', '=(B3-B2)/B4')
    ws1.write('A6', 'True I =')
    ws1.write('B6', 1.640533)

    headers1 = ["i", "xi", "f(xi)", "Trap. Multiplier", "Simpson's Mult.", "Notes"]
    for c, h in enumerate(headers1):
        ws1.write(7, c, h, header_fmt)

    for i in range(11):
        row = i + 8
        ws1.write(row, 0, i)
        ws1.write_formula(row, 1, f'=$B$2+{i}*$B$5')
        # f(x) = 0.2 + 25x - 200x^2 + 675x^3 - 900x^4 + 400x^5
        ws1.write_formula(row, 2, f'=0.2+25*B{row+1}-200*B{row+1}^2+675*B{row+1}^3-900*B{row+1}^4+400*B{row+1}^5', num_fmt)
        
        # Trap Multiplier
        if i == 0 or i == 10: ws1.write(row, 3, 1)
        else: ws1.write(row, 3, 2)
        
        # Simpson Multiplier (1, 4, 2, 4, ..., 4, 1)
        if i == 0 or i == 10: ws1.write(row, 4, 1)
        elif i % 2 == 1: ws1.write(row, 4, 4)
        else: ws1.write(row, 4, 2)

    # Integration Results
    ws1.write('A20', 'Trapezoidal Result (I)', bold_fmt)
    ws1.write_formula('C20', '=(B5/2)*SUMPRODUCT(C9:C19, D9:D19)', num_fmt)
    
    ws1.write('A21', 'Simpson\'s 1/3 Result (I)', bold_fmt)
    ws1.write_formula('C21', '=(B5/3)*SUMPRODUCT(C9:C19, E9:E19)', num_fmt)
    
    ws1.write('A22', 'εt (Trapezoidal)', bold_fmt)
    ws1.write_formula('C22', '=ABS(($B$6-C20)/$B$6)', pct_fmt)
    
    ws1.write('A23', 'εt (Simpson\'s)', bold_fmt)
    ws1.write_formula('C23', '=ABS(($B$6-C21)/$B$6)', pct_fmt)

    # --- P2_RK4 ---
    ws2 = workbook.add_worksheet('P2_RK4')
    ws2.set_column('A:H', 15)
    
    ws2.write('A1', '4th-Order Runge-Kutta (RK4)', bold_fmt)
    ws2.write('A2', 'h =')
    ws2.write('B2', 0.5)
    ws2.write('A3', 'x0 =')
    ws2.write('B3', 0)
    ws2.write('A4', 'y0 =')
    ws2.write('B4', 1)

    headers2 = ["Step (i)", "xi", "yi", "k1", "k2", "k3", "k4", "yi+1"]
    for c, h in enumerate(headers2):
        ws2.write(6, c, h, header_fmt)

    # Step 0
    ws2.write(7, 0, 0)
    ws2.write(7, 1, 0)
    ws2.write(7, 2, 1)
    
    # ODE: f(x, y) = -2x^3 + 12x^2 - 20x + 8.5
    for i in range(4):
        row = i + 7
        # k1 = f(xi)
        ws2.write_formula(row, 3, f'=-2*B{row+1}^3+12*B{row+1}^2-20*B{row+1}+8.5', num_fmt)
        # k2 = f(xi + 0.5h)
        ws2.write_formula(row, 4, f'=-2*(B{row+1}+0.5*$B$2)^3+12*(B{row+1}+0.5*$B$2)^2-20*(B{row+1}+0.5*$B$2)+8.5', num_fmt)
        # k3 = f(xi + 0.5h)
        ws2.write_formula(row, 5, f'=-2*(B{row+1}+0.5*$B$2)^3+12*(B{row+1}+0.5*$B$2)^2-20*(B{row+1}+0.5*$B$2)+8.5', num_fmt)
        # k4 = f(xi + h)
        ws2.write_formula(row, 6, f'=-2*(B{row+1}+$B$2)^3+12*(B{row+1}+$B$2)^2-20*(B{row+1}+$B$2)+8.5', num_fmt)
        # yi+1 = yi + (h/6)(k1+2k2+2k3+k4)
        ws2.write_formula(row, 7, f'=C{row+1}+($B$2/6)*(D{row+1}+2*E{row+1}+2*F{row+1}+G{row+1})', num_fmt)
        
        # Next row xi and yi
        if i < 3:
            ws2.write(row + 1, 0, i + 1)
            ws2.write_formula(row + 1, 1, f'=B{row+1}+$B$2')
            ws2.write_formula(row + 1, 2, f'=H{row+1}')
    
    # Final row xi
    ws2.write(11, 0, 4)
    ws2.write_formula(11, 1, f'=B11+$B$2')
    ws2.write(11, 2, '—', border_fmt)

    # Chart
    chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth_with_markers'})
    chart.add_series({
        'name':       'RK4 Solution',
        'categories': '=P2_RK4!$B$8:$B$12',
        'values':     '=P2_RK4!$C$8:$C$12',
    })
    chart.set_title({'name': 'RK4 Numerical Solution'})
    chart.set_x_axis({'name': 'x'})
    chart.set_y_axis({'name': 'y'})
    ws2.insert_chart('J2', chart)

    workbook.close()

if __name__ == '__main__':
    create_final_answer_key()

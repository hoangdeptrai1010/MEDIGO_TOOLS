import openpyxl, sys

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('baocaokpi_thang8_hoanthien.xlsx', data_only=False)
ws_ds = wb['kpi dược sĩ']

for r in range(3, ws_ds.max_row + 1):
    if ws_ds.cell(r, 2).value:
        ws_ds.cell(r, 6, f"=0.25*I{r}")
        # Standard Excel compatible formula using IF
        f_k = f'=_xlfn.IFERROR(_xlfn.XLOOKUP(B{r}, \'kpi nhà thuốc\'!A:A, \'kpi nhà thuốc\'!G:G), IF(AND(OR(C{r}="DSXC", C{r}="DSCD", C{r}="DSBC", C{r}="DSTV"), G{r}>=F{r}), AD{r} * _xlfn.IFS(AND(I{r}>=T{r}, E{r}>=D{r}), 0.015, I{r}>=T{r}, 0.013, AND(I{r}>=S{r}, E{r}>=D{r}), 0.013, I{r}>=S{r}, 0.011, AND(I{r}>=R{r}, E{r}>=D{r}), 0.012, I{r}>=R{r}, 0.010, TRUE, 0), 0))'
        ws_ds.cell(r, 11, f_k)

wb.save('baocaokpi_thang8_hoanthien.xlsx')
wb.close()
print("Updated formulas in baocaokpi_thang8_hoanthien.xlsx")

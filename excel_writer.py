
def write_to_excel(workbook, worksheet, outputs):
        bold = workbook.add_format({'bold': True})
        bold_italic = workbook.add_format({'bold': True, 'italic':True})
        border_bold = workbook.add_format({'border':True,'bold':True})
        border_bold_grey = workbook.add_format({'border':True,'bold':True,'bg_color':'#d3d3d3'})
        border = workbook.add_format({'border':True})
        worksheet.set_column('B:D', 22)
        worksheet.set_column('E:F', 33)
        row = 0
        col = 0
        worksheet.write(row,col,'Sl No',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Listing Type',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Sales Price',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Rent',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Sold Amt',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Str #',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Street Name',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'City',border_bold_grey)
        col = col + 1
        # worksheet.write(row,col,'State',border_bold_grey)
        # col = col + 1
        worksheet.write(row,col,'Zip',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Home type',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'BR',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'BA',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Sq ft',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Zestimate',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Rent Zestimate',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Lot Size',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Days on Zillow',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Link to Zillow',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'MLS',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Views',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'HOA',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Shopper Saves',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Nearby Elementary',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Nearby Middle',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Nearby High',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Tax Year_1',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Property Taxes_1',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Tax Assessment_1',border_bold_grey)
       	col = col + 1
       	worksheet.write(row,col,'Tax Year_2',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Property Taxes_2',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Tax Assessment_2',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Date_1',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Event_1',border_bold_grey)
       	col = col + 1
       	worksheet.write(row,col,'Price_1',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Source_1',border_bold_grey)
       	col = col + 1
       	worksheet.write(row,col,'Date_2',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Event_2',border_bold_grey)
       	col = col + 1
       	worksheet.write(row,col,'Price_2',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Source_2',border_bold_grey)
       	col = col + 1
       	worksheet.write(row,col,'Date_3',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Event_3',border_bold_grey)
       	col = col + 1
       	worksheet.write(row,col,'Price_3',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Source_3',border_bold_grey)
       	row = row + 1
        i = 0

        for output in outputs:
            i = i + 1
            col = 0
            worksheet.write(row, col, i, border)
            col = col + 1
            worksheet.write(row, col, output["Listing Type"] if 'Listing Type' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Sales Price"] if 'Sales Price' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Rent"] if 'Rent' in output else '', border)
            col = col + 1
            worksheet.write(row, col, output["Sold Amt"] if 'Sold Amt' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Str #"] if 'Str #' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Street Name"] if 'Street Name' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["City"] if 'City' in output else '',border)
            col = col + 1
            # worksheet.write(row, col, output["State"] if output.has_key('State') else '',border)
            # col = col + 1
            worksheet.write(row, col, output["Zip"] if 'Zip' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Home type"] if 'Home type' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["BR"] if 'BR' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["BA"] if 'BA' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Sq ft"] if 'Sq ft' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Zestimate"] if 'Zestimate' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Rent Zestimate"] if 'Rent Zestimate' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Lot Size"] if 'Lot Size' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Days on Zillow"] if 'Days on Zillow' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Link to Zillow"] if 'Link to Zillow' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["MLS"] if 'MLS' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Views"] if 'Views' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["HOA"] if 'HOA' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Shopper Saves"] if 'Shopper Saves' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Nearby Elementary"] if 'Nearby Elementary' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Nearby Middle"] if 'Nearby Middle' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Nearby High"] if 'Nearby High' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Tax Year_1"] if 'Tax Year_1' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Property Taxes_1"] if 'Property Taxes_1' in output else '',border)
            col = col + 1
            worksheet.write(row, col,  output["Tax Assessment_1"] if 'Tax Assessment_1' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Tax Year_2"] if 'Tax Year_2' in output else '', border)
            col = col + 1
            worksheet.write(row, col, output["Property Taxes_2"] if 'Property Taxes_2' in output else '',border)
            col = col + 1
            worksheet.write(row, col,  output["Tax Assessment_2"] if 'Tax Assessment_2' in output else '',border)
            col = col + 1
            worksheet.write(row, col,  output["Date_1"] if 'Date_1' in output else '',border)
            col = col + 1
            worksheet.write(row, col,  output["Event_1"] if 'Event_1' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Price_1"] if 'Price_1' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Source_1"] if 'Source_1' in output else '',border)
            col = col + 1
            worksheet.write(row, col,  output["Date_2"] if 'Date_2' in output else '',border)
            col = col + 1
            worksheet.write(row, col,  output["Event_2"] if 'Event_2' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Price_2"] if 'Price_2' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Source_2"] if 'Source_2' in output else '',border)
            col = col + 1
            worksheet.write(row, col,  output["Date_3"] if 'Date_3' in output else '',border)
            col = col + 1
            worksheet.write(row, col,  output["Event_3"] if 'Event_3' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Price_3"] if 'Price_3' in output else '',border)
            col = col + 1
            worksheet.write(row, col, output["Source_3"] if 'Source_3' in output else '',border)
            row = row + 1

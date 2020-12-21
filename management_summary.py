import pandas as pd
from tkinter.filedialog import askopenfilename
from collections import defaultdict

learnpro_runtime = input("when was this data pulled from learnpro? (format = dd-mm-yy)")
learnpro_date = pd.to_datetime(learnpro_runtime, format='%d-%m-%y')

def mh():
    filename = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/Named Lists HSE',
                               title="Choose the relevant moving and handling file"
                               )
    df = pd.read_excel(filename, sheet_name='data')
    print(df.columns)

    piv1 = pd.pivot_table(df[((df['Scope'] == 1) & (df['Area']!='Board-Wide Services'))], index=['Area','Sector/Directorate/HSCP'], columns='Assessed Category',
                          values='Scope', aggfunc='count', margins=True, margins_name='Total')
    piv1['Compliant'] = piv1[['1','2','3 or more']].sum(axis=1)
    piv1['Compliance %'] = ((piv1[['Compliant', 'Not at work ≥ 28 days']].sum(axis=1) / piv1['Total']) * 100).round(2)

    piv2 = pd.pivot_table(df[(df['Scope'] == 1)],
                          index=['Area', 'Sector/Directorate/HSCP', 'department'], columns='Assessed Category',
                          values='Scope', aggfunc='count', margins=True, margins_name='Headcount')
    piv2['Compliant'] = piv2[['1','2','3 or more']].sum(axis=1)
    piv2['Compliant'].fillna(0, inplace=True)
    piv2['Compliance %'] = ((piv2[['Compliant', 'Not at work ≥ 28 days']].sum(axis=1)/ piv2['Headcount']) * 100).round(1)
    piv2['Size of Dept'] = pd.cut(piv2['Headcount'], bins=[0,15, 30, 60, 10000], labels=['Small', 'Medium', 'Large', 'Extra Large'])

    phase_lookup = pd.read_excel('W:/Learnpro/phase_lookup-MH.xlsx')
    phase_lookup = defaultdict(lambda:"Phase 3", dict(zip(phase_lookup['department'], phase_lookup['phase'])))
    print(phase_lookup)
    piv2 = piv2.reset_index()
    piv2['Phase'] = piv2['department'].map(phase_lookup)
    piv2 = piv2[['Sector/Directorate/HSCP', 'department', 'Phase', 'Headcount','Not at work ≥ 28 days', 'Compliant', 'Size of Dept',
                 'Compliance %']]
    piv2.sort_values([ 'Size of Dept', 'Compliant'], ascending=[False, True], inplace=True)
    piv2.drop(piv2.tail(1).index, inplace=True)
    return piv1, piv2

def falls():
    filename = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/Named Lists HSE',
                               title="Choose the relevant Falls file"
                               )
    df = pd.read_excel(filename, sheet_name='Export')
    df['Headcount'] = 1
    exclusions = ['Centre For Population Health', 'Board Administration', 'Corporate Communications',
                  'COVID-19 Temp Nurses', 'eHealth', 'Estates and Facilities', 'Finance', 'HR and OD',
                  'Non Paid Employees', 'Public Health']
    piv1 = pd.pivot_table(df[(df['Falls Compliant'] != 'Out of scope') & (~df['Sector/Directorate/HSCP'].isin(exclusions))],
                          index=['Area', 'Sector/Directorate/HSCP'], columns='Falls Compliant',
                          values='Headcount', aggfunc='count', margins=True, margins_name='Total', fill_value=0)
    piv1['Compliance %'] = ((piv1[['Not at work ≥ 28 days','Complete']].sum(axis=1)) / piv1['Total'] * 100).round(1)

    piv2 = pd.pivot_table(df[(df['Falls Compliant'] != 'Out of scope')],
                          index=['Area', 'Sector/Directorate/HSCP', 'department'], columns='Falls Compliant',
                          values='Headcount', aggfunc='count', margins=True, margins_name='Total', fill_value=0)

    piv2['Compliant'] = (piv2[['Not at work ≥ 28 days','Complete']].sum(axis=1))
    piv2['Compliance %'] = ((piv2['Compliant'] / piv2['Total']) * 100).round(1)
    piv2['Size of Dept'] = piv2['Size of Dept'] = pd.cut(piv2['Total'], bins=[0,15, 30, 60, 10000],
                                                         labels=['Small', 'Medium', 'Large', 'Extra Large'])
    phase_lookup = pd.read_excel('W:/Learnpro/phase_lookup-Falls.xlsx')
    phase_lookup = defaultdict(lambda: "Phase 3", dict(zip(phase_lookup['department'], phase_lookup['Phase'])))
    piv2 = piv2.reset_index()
    piv2['Phase'] = piv2['department'].map(phase_lookup)
    piv2 = piv2[['Area', 'Sector/Directorate/HSCP', 'department', 'Complete', 'Not at work ≥ 28 days', 'No Account',
                 'Not Complete', 'Not Undertaken', 'Total', 'Phase', 'Size of Dept', 'Compliance %']]
    piv2.sort_values(['Size of Dept', 'Complete'], ascending=[False, True], inplace=True)
    piv2.drop(piv2.tail(1).index, inplace=True)
    return piv1, piv2


def sharps():
    filename = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/HSE Sharps and Skins',
                               title="Choose the relevant Sharps file"
                               )
    df = pd.read_excel(filename, sheet_name='Export')
    df['Headcount'] = 1
    dept_notatwork = {}
    sector_notatwork = {}
    sector_complete = {}
    dept_complete = {}

    for i in df['department'].unique().tolist():
        dfx = df[(df['department'] == i) & ((df['GGC Module'] == 'Not at work ≥ 28 days') |
                                            (df['NES Module'] == 'Not at work ≥ 28 days'))]
        dfy = df[(df['department'] == i) & (df['GGC Module'] == 'Complete') &
                 (df['NES Module'].isin(['Complete', 'Out Of Scope']))]
        dept_complete[i] = len(dfy)
        dept_notatwork[i] = len(dfx)
    for i in df['Sector/Directorate/HSCP'].unique().tolist():
        dfx = df[(df['Sector/Directorate/HSCP'] == i) & ((df['GGC Module'] == 'Not at work ≥ 28 days') |
                                                         (df['NES Module'] == 'Not at work ≥ 28 days'))]
        dfy = df[(df['Sector/Directorate/HSCP'] == i) & (df['GGC Module'] == 'Complete') &
                 (df['NES Module'].isin(['Complete', 'Out Of Scope']))]
        sector_notatwork[i] = len(dfx)
        sector_complete[i] = len(dfy)


    piv1 = pd.pivot_table(df[(df['GGC Module']!='Out Of Scope')], index=['Area', 'Sector/Directorate/HSCP'],
                          columns='Compliant', values='Headcount', aggfunc='count', margins=True,
                          margins_name='Total', fill_value=0)
    piv1 = piv1.reset_index()
    piv1['Not at work ≥ 28 days'] = piv1['Sector/Directorate/HSCP'].map(sector_notatwork)
    piv1['Complete'] = piv1['Sector/Directorate/HSCP'].map(sector_complete)

    print(piv1.columns)
    piv1.rename(inplace=True, columns={0:'Not Compliant', 1:'Compliant'})

    # deleted below for double counting
    #piv1['Not Compliant'] = piv1['Not Compliant'] - piv1['Not at work ≥ 28 days']
    piv1['Compliance %'] = (piv1[['Complete', 'Not at work ≥ 28 days']].sum(axis=1) / piv1['Total'] * 100).round(1)
    piv1 = piv1[['Area', 'Sector/Directorate/HSCP', 'Not Compliant', 'Complete', 'Not at work ≥ 28 days', 'Total',
                 'Compliance %']]
    piv1.drop(piv1.tail(1).index, inplace=True)
    total_row = ['Total',' ',piv1['Not Compliant'].sum(),piv1['Complete'].sum(),
                              piv1['Not at work ≥ 28 days'].sum(),piv1['Total'].sum(),
                 (((piv1['Complete'].sum() + piv1['Not at work ≥ 28 days'].sum()) / piv1['Total'].sum()) * 100).round(1)]

    piv1.loc[len(piv1)] = total_row

    piv2 = pd.pivot_table(df[(df['GGC Module'] != 'Out Of Scope')],
                          index=['Area', 'Sector/Directorate/HSCP', 'department'],
                          columns='Compliant', values='Headcount', aggfunc='count', margins=True,
                          margins_name='Total', fill_value=0)
    piv2 = piv2.reset_index()
    piv2['Size of Dept'] = pd.cut(piv2['Total'], bins=[0, 15, 30, 60, 10000],
                                  labels=['Small', 'Medium', 'Large', 'Extra Large'])
    phase_lookup = pd.read_excel('W:/Learnpro/phase_lookup-sharps.xlsx')
    phase_lookup = defaultdict(lambda: "Phase 3", dict(zip(phase_lookup['department'], phase_lookup['phase'])))
    piv2['Phase'] = piv2['department'].map(phase_lookup)
    piv2['Not at work ≥ 28 days'] = piv2['department'].map(dept_notatwork)
    piv2.rename(inplace=True, columns={0:'Not Compliant', 1:'Compliant'})
    piv2['Not Compliant'] = piv2['Not Compliant']
    piv2['Complete'] = piv2['department'].map(dept_complete)
    piv2['Compliance %'] = ((piv2[['Complete', 'Not at work ≥ 28 days']].sum(axis=1) / piv2['Total']) * 100).round(1)
    piv2.sort_values(['Size of Dept', 'Complete'], ascending=[False, True], inplace=True)
    piv2.drop(piv2.tail(1).index, inplace=True)
    piv2 = piv2[['Area', 'Sector/Directorate/HSCP', 'department', 'Phase', 'Not Compliant', 'Not at work ≥ 28 days',
                 'Complete', 'Total', 'Size of Dept', 'Compliance %']]

    return piv1, piv2




def build_file():

    with pd.ExcelWriter('W:/Learnpro/Named Lists HSE/' + learnpro_date.strftime('%Y-%m-%d') + '/' + learnpro_date.strftime(
        '%Y%m%d') + ' - HSE Management Summary.xlsx', engine='xlsxwriter') as writer:
        #Cover page
        workbook = writer.book

        worksheet = workbook.add_worksheet('Cover')
        worksheet.hide_gridlines(2)
        worksheet.set_column('C:D', 26)
        worksheet.set_column('B:B', 36)
        worksheet.set_column('A:A', 5)

        # formats
        header_format = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 16})
        subheader = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 14})
        subheadernoBold = workbook.add_format({'font_name': 'Arial', 'font_size': 14})
        regularBold = workbook.add_format({'bold': True})
        table_format_ul = workbook.add_format({'bg_color':'#005EB8', 'font_color':'white', 'underline':True})
        table_format = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'white'})
        percent_highlight = workbook.add_format({'bg_color':'#005EB8', 'font_color':'white', 'num_format':'0.0%'})
        merge_format = workbook.add_format({'bold':1, 'border':1, 'align':'center'})
        wrap_format = workbook.add_format({'text_wrap':1,'align':'vcenter', 'bg_color':'#005EB8',
                                                 'font_color':'white'})
        back_button = table_format_ul = workbook.add_format({'bg_color':'#005EB8', 'font_color':'white',
                                                             'underline':True, 'align':'center', 'valign':'vcenter'})

        valign_email_link = workbook.add_format({'align':'vcenter', 'bold':1, 'bg_color':'#005EB8',
                                                 'font_color':'white', 'underline':1})
        # write headers
        worksheet.write(0, 1, 'Health & Safety Training Summary Report', header_format)
        worksheet.write(1, 1, 'Date:', subheader)
        worksheet.write(1, 2, f'{learnpro_date.strftime("%d %B %Y")}', subheadernoBold)
        worksheet.insert_image('D1', 'W:/Danny/ggclogo.jpg', {'x_scale':0.25, 'y_scale':0.25})

        # write table
        worksheet.write_row('B4', ['HSE Area', 'Sector Level', 'Department Level'], header_format)
        worksheet.write_row('B5', ['Falls', 'Sector/Directorate/HSCP', 'Departmental List'], table_format)
        worksheet.write_row('B6', ['Sharps', 'Sector/Directorate/HSCP', 'Departmental List'], table_format)
        worksheet.write_row('B7', ['Moving & Handling', 'Sector/Directorate/HSCP', 'Departmental List'], table_format)
        worksheet.write_url('C6', "internal:'Sharps-SectorDirectorate'!A1",
                            string='Sector/Dir/HSCP',cell_format=table_format_ul)
        worksheet.write_url('D6', "internal:'Sharps-ByDepartment'!A1",
                            string='Departmental List', cell_format=table_format_ul)
        worksheet.write_url('C5', "internal:'Falls-SectorDirectorate'!A1",
                            string='Sector/Dir/HSCP', cell_format=table_format_ul)
        worksheet.write_url('D5', "internal:'Falls-ByDepartment'!A1",
                            string='Departmental List', cell_format=table_format_ul)
        worksheet.write_url('C7', "internal:'MovingHandling-SectorDirectorat'!A1",
                            string='Sector/Dir/HSCP', cell_format=table_format_ul)
        worksheet.write_url('D7', "internal:'MovingHandling-Department'!A1",
                            string='Departmental List', cell_format=table_format_ul)

        # Explanation of depts
        worksheet.write('B13', 'Department Size', subheader)
        worksheet.write(13, 1, 'Department lists have a column indicating size, where:', regularBold)
        worksheet.write(14, 1, '- Small is <15', wrap_format)
        worksheet.write(15, 1, '- Medium is 15-29', wrap_format)
        worksheet.write(14, 2, '- Large is 30-60', wrap_format)
        worksheet.write(15, 2, '- Extra Large is 60+', wrap_format)
        # Guidance document
        worksheet.write('B9', 'Contact', subheader)
        worksheet.write('B10', 'For clarification about training eligibility or other HSE queries, ' +
                                   'please email:', wrap_format )
        worksheet.write_url('C10', 'mailto:Cameron.Raeburn@ggc.scot.nhs.uk',
                            string='Cameron Raeburn', cell_format=valign_email_link)
        worksheet.write('B11', 'If you believe someone is shown as in or out of scope erroneously' +
                             ' or for development queries, please email:', wrap_format)
        worksheet.write_url('C11', 'mailto:Daniel.Furnivall@ggc.scot.nhs.uk',
                      string=' Daniel Furnivall', cell_format=valign_email_link)




        # Sharps pivots
        s1.to_excel(writer, startcol=0, startrow=1, sheet_name='Sharps-SectorDirectorate', index=False)
        s2.to_excel(writer, startcol=0, startrow=2, sheet_name='Sharps-ByDepartment', index=False)
        # Sharps sector titles
        sharpsSectorTitle = writer.sheets['Sharps-SectorDirectorate']
        sharpsSectorTitle.conditional_format(f'G3:G{len(s1) + 2}', {'type': '3_color_scale'})
        sharpsSectorTitle.write_string(0,0, 'Sharps - Sector / Directorate Summary', subheader)
        sharpsSectorTitle.write_url('G1', "internal:'Cover'!A1",
                            string='Back', cell_format=back_button)
        # Sharps dept titles
        sharpsDeptTitle = writer.sheets['Sharps-ByDepartment']
        sharpsDeptTitle.conditional_format(f'J4:J{len(s2) + 3}', {'type': '3_color_scale'})
        sharpsDeptTitle.write_string(0, 0, 'Sharps - Department Summary', subheader)
        sharpsDeptTitle.autofilter(f'A3:J{len(f2) + 3}')
        sharpsDeptTitle.freeze_panes(3, 0)
        sharpsDeptTitle.write_url('G1', "internal:'Cover'!A1",
                            string='Back', cell_format=back_button)
        # additional sharps bits
        sharpsDeptTitle.write_string('L1', 'Compliance %')
        sharpsDeptTitle.write_string('M1', 'Completed')
        sharpsDeptTitle.write_string('N1', 'Not at work ≥ 28 days')
        sharpsDeptTitle.write_string('O1', 'Compliant')
        sharpsDeptTitle.write_string('P1', 'Headcount')
        sharpsDeptTitle.write_formula('L2', '=O2/P2', percent_highlight)
        sharpsDeptTitle.write_formula('M2', '=SUBTOTAL(109,G:G)')
        sharpsDeptTitle.write_formula('N2', '=SUBTOTAL(109,F:F)')
        sharpsDeptTitle.write_formula('O2', '=M2+N2')
        sharpsDeptTitle.write_formula('P2', '=SUBTOTAL(109,H:H)')

        # Falls pivots
        f1.to_excel(writer, startcol=0, startrow=1, sheet_name='Falls-SectorDirectorate')
        f2.to_excel(writer, startcol=0, startrow=2, sheet_name='Falls-ByDepartment', index=False)
        # Falls sector titles
        fallsSectorTitle = writer.sheets['Falls-SectorDirectorate']
        fallsSectorTitle.conditional_format(f'I3:I{len(f1) + 2}', {'type': '3_color_scale'})
        fallsSectorTitle.write_string(0,0, 'Falls - Sector / Directorate Summary', subheader)
        fallsSectorTitle.write_url('G1', "internal:'Cover'!A1",
                            string='Back', cell_format=back_button)
        # Falls titles - dept
        fallsDeptTitle = writer.sheets['Falls-ByDepartment']
        fallsDeptTitle.conditional_format(f'L4:L{len(f2) + 3}', {'type': '3_color_scale'})
        fallsDeptTitle.autofilter(f'A3:L{len(f2)+3}')
        fallsDeptTitle.write_string(0,0, 'Falls - Department Summary', header_format)
        fallsDeptTitle.freeze_panes(3, 0)
        fallsDeptTitle.write_url('G1', "internal:'Cover'!A1",
                            string='Back', cell_format=back_button)
        # Additional falls dept bits
        fallsDeptTitle.write_string('N1', 'Compliance %')
        fallsDeptTitle.write_string('O1', 'Complete')
        fallsDeptTitle.write_string('P1', 'Not at work ≥ 28 days')
        fallsDeptTitle.write_string('Q1', 'Compliant')
        fallsDeptTitle.write_string('R1', 'Headcount')
        fallsDeptTitle.write_formula('N2', '=Q2/R2', percent_highlight)
        fallsDeptTitle.write_formula('O2', '=SUBTOTAL(109,D:D)')
        fallsDeptTitle.write_formula('P2', '=SUBTOTAL(109,E:E)')
        fallsDeptTitle.write_formula('Q2', '=O2+P2')
        fallsDeptTitle.write_formula('R2', '=SUBTOTAL(109,I:I)')


        # MH pivots
        mh1.to_excel(writer, startcol=0, startrow=2, sheet_name='MovingHandling-SectorDirectorat')
        mh2.to_excel(writer, startcol=0, startrow=2, sheet_name='MovingHandling-Department', index=False)
        # MH Titles - sector sheet
        mhSectorTitle = writer.sheets['MovingHandling-SectorDirectorat']
        mhSectorTitle.conditional_format(f'J3:J{len(mh1)+3}', {'type':'3_color_scale'})
        mhSectorTitle.write_string(0, 0, 'Moving and Handling Assessments - Sector / Directorate Summary', subheader)
        mhSectorTitle.merge_range('C2:F2', 'No. of Assessments', merge_format)
        mhSectorTitle.write_url('K1', "internal:'Cover'!A1",
                            string='Back', cell_format=back_button)

        # MH Titles - dept sheet
        mhDeptTitle = writer.sheets['MovingHandling-Department']
        mhDeptTitle.conditional_format(f'H4:H{len(mh2) + 3}', {'type':'3_color_scale'})
        mhDeptTitle.autofilter(f'A3:H{len(f2) + 3}')
        mhDeptTitle.write_string(0, 0, 'Moving and Handling - Department Summary', subheader)
        mhDeptTitle.freeze_panes(3, 0)
        mhDeptTitle.write_url('H1', "internal:'Cover'!A1",
                            string='Back', cell_format=back_button)
        # Additional dept sheet bits

        mhDeptTitle.write_string('I1', 'Compliance %')
        mhDeptTitle.write_string('J1', 'Compliant')
        mhDeptTitle.write_string('K1', 'Not at work ≥ 28 days')
        mhDeptTitle.write_string('L1', 'Completed')
        mhDeptTitle.write_string('M1', 'Headcount')
        mhDeptTitle.write_formula('I2', '=J2/M2', percent_highlight)
        mhDeptTitle.write_formula('J2', '=L2+K2')
        mhDeptTitle.write_formula('K2', '=SUBTOTAL(109,E:E)')
        mhDeptTitle.write_formula('L2', '=SUBTOTAL(109,F:F)')
        mhDeptTitle.write_formula('M2', '=SUBTOTAL(109,D:D)')
    # Get the xlsxwriter workbook and worksheet objects.

        worksheet = writer.sheets['Sharps-SectorDirectorate']
        worksheet.conditional_format(f'G3:G{len(s1) + 2}', {'type': '3_color_scale'})




s1, s2 = sharps()

f1, f2 = falls()
mh1, mh2 = mh()

build_file()
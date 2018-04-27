from flask import Flask, render_template, request, send_file, make_response, Response, url_for,redirect, abort, Response
from flask_restful import abort
import xlsxwriter
from io import BytesIO
import pandas as pd
import numpy as np
from json import dumps
from math import radians, cos, sin, asin, sqrt
app=Flask(__name__)

@app.route('/home_final')
def home_final():
	return render_template("home_final.html")
@app.route('/index')
def index():
	return render_template("index.html")	
@app.route('/pci_planner')
def pci_planner():
	return render_template("pci_planner.html")
@app.route('/aboutme')
def aboutme():
	return render_template("aboutme.html")	
@app.route('/contactme')
def contactme():
	return render_template("contactme.html")
@app.route('/pci_reuse')
def pci_reuse():
	return render_template("pci_reuse.html")
@app.route('/audit_tdd')	
def audit_tdd():
	return render_template("audit_tdd.html")
@app.route('/audit_fdd')	
def audit_fdd():
	return render_template("audit_fdd.html")


@app.route('/login', methods=['GET', 'POST'])

def login():
	error = None
	if request.method == 'POST':
		if request.form['username'] != 'admin' or request.form['password'] != 'admin':
			return redirect(url_for('error'))
		else:
			return redirect(url_for('home_final'))
	return render_template('login.html', error=error)

@app.route('/error', methods=['GET', 'POST'])

def error():
	error = None
	if request.method == 'POST':
		if request.form['username'] != 'admin' or request.form['password'] != 'admin':
			error = 'Invalid Credentials. Please try again.'
		else:
			return redirect(url_for('home_final'))
	return render_template('error.html', error=error)

@app.route('/handle_data', methods=['POST','GET'])

def handle_data():

	reuse=request.form['reuse_distance']
	inputdata=request.files['new_site_data']
	rfdata=request.files['rf_database']
	df=pd.read_excel(inputdata)
	df100=pd.read_excel(rfdata)
	
	

	def popuperror():  

		error_message = dumps({'Message': 'REDUCE THE RE-USE DISTANCE AND RUN THE PROGRAM AGAIN'})
		abort(Response(error_message, 401))
		

	def haversine(lon1, lat1, lon2, lat2):
			"""
			Calculate the great circle distance between two points 
			on the earth (specified in decimal degrees)
			"""
			# convert decimal degrees to radians 
			lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])
			# haversine formula 
			dlon = lon2 - lon1 
			dlat = lat2 - lat1 
			a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
			c = 2 * asin(sqrt(a)) 
			# Radius of earth in kilometers is 6371
			km = 6371* c
			return km
	dist_1=[]
	pci_1=[]
	pci_final=[]
	lat_new_site=[]
	long_new_site=[]
	site_name=[]
	cell_name=[]
	Reuse_dist=[]
	min_reuse_dist=float(reuse)
	inp=df
				
	df_n_1=df100
		
	for i in range(len(inp['site'])):

		lat_1= inp.iloc[i,1]
		lon_1= inp.iloc[i,2]    
		site_=inp.iloc[i,0]
		#print("details of the new site \n ", site_,lat_1,lon_1)
		df_n_1.sort_values(['Cell'])
		#print("tail of df_n_1\n", df_n_1.tail())
		#print("source lat long" , lat_1,lon_1)
		for j in range(len(df_n_1['Cell'])):
		
	
			pci_1.append(df_n_1.iloc[j,3])
			lon_2=df_n_1.iloc[j,2].astype(float)
			lat_2=df_n_1.iloc[j,1].astype(float)
			d=haversine(lon_1,lat_1,lon_2,lat_2)
			dist_1.append(d)
   
		df_n_2=pd.DataFrame(list(zip(pci_1,dist_1)))
		#print(df_n_2)
		df_n_2.columns=('PCI','Distance_from_new_site') 
		#df_n_2=df_n_2.sort_values(['Distance_from_new_site'], ascending=False)
		
		#df_n_2.to_excel('test123456.xlsx')

		df_n_2=df_n_2.groupby('PCI')['Distance_from_new_site'].agg(['min']).reset_index()
		df_n_2.columns=('PCI','Distance_from_new_site')
		#print(df_n_2)
		xz=df_n_2['Distance_from_new_site']<min_reuse_dist
		
		df_n_3=df_n_2[xz]
		if df_n_3.empty:
		   popuperror()
		df_n_3=df_n_3.sort_values(['Distance_from_new_site'], ascending= True)

		#print(df_n_2)
			#button = Button(root, text = "show message box", command = self.popuperror)
		
				
			
		df_n_4=pd.concat([df_n_2,df_n_3])
		df_n_4=df_n_4.drop_duplicates(subset='PCI',keep=False).sort_values(['Distance_from_new_site'], ascending=False)
		if df_n_4.empty:
		   popuperror()
				
		#print("PCIs that can be used \n", df_n_4)
		p1=int(df_n_4.iloc[0,0]/3)*3
		p2=p1+1
		p3=p1+2
		site_name.append(site_)
		site_name.append(site_) 
		site_name.append(site_)
		cell_name.append(site_+'A')
		cell_name.append(site_+'B')
		cell_name.append(site_+'C')
		lat_new_site.append(lat_1)
		lat_new_site.append(lat_1)
		lat_new_site.append(lat_1)
		long_new_site.append(lon_1)
		long_new_site.append(lon_1)
		long_new_site.append(lon_1)
		pci_final.append(p1)
		pci_final.append(p2)
		pci_final.append(p3)
		Reuse_dist.append(df_n_4.iloc[0,1])
		Reuse_dist.append(df_n_4.iloc[0,1])
		Reuse_dist.append(df_n_4.iloc[0,1])                
		df_n_5=pd.DataFrame(list(zip(site_name,cell_name,lat_new_site,long_new_site,pci_final,Reuse_dist)))
		df_n_5.columns=('site_name','cell_name','lat_new_site','long_new_site','pci_final','Reuse_dist')
		df_n_6=df_n_5.iloc[:,1:5]
		df_n_6.columns=df_n_1.columns

	output = BytesIO()
	writer = pd.ExcelWriter(output, engine='xlsxwriter')
	workbook = writer.book

	format = workbook.add_format()
	format.set_bg_color('#eeeeee')
	


	df_n_3.to_excel(writer, startrow = 0, merge_cells = False, sheet_name = "Sheet_1")        
	df_n_4.to_excel(writer, startrow = 0, merge_cells = False, sheet_name = "Sheet_2")
	df_n_5.to_excel(writer, startrow=1, merge_cells=False, sheet_name="FINAL PCI Design")


	
	
	writer.save()
	writer.close()
	output.seek(0)  
	#return Response(df1,headers={"Content-disposition":"attachment; filename=myplot.xls"}) 
	return send_file(output, attachment_filename="testing.xlsx", as_attachment=True)
	

@app.route('/reuse_calc', methods=['POST','GET'])

def reuse_calc():


	def haversine(lon1, lat1, lon2, lat2):
			"""
			Calculate the great circle distance between two points 
			on the earth (specified in decimal degrees)
			"""
			# convert decimal degrees to radians 
			lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])
			# haversine formula 
			dlon = lon2 - lon1 
			dlat = lat2 - lat1 
			a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
			c = 2 * asin(sqrt(a)) 
			# Radius of earth in kilometers is 6371
			km = 6371* c
			return km

	rf_data_reuse=request.files['new_site_data_1']
	df_99=pd.read_excel(rf_data_reuse)
	tt=pd.DataFrame(df_99['PCI'].unique(), columns=['PCI'])
	tt['reuse']=(df_99.groupby('PCI')['PCI'].transform('count')-1)
	#print((tt))


	pci_list=list(df_99['PCI'].unique())
	#print(pci_list)
	source_cell=[]
	target_cell=[]
	distance=[]
	pci_final=[]
	#pci_avail=np.arange(0,504)
	#print(pci_avail)
#pci_list=[441]
	for p in pci_list:
		
		yyy=df_99['PCI']==p
		temp=df_99[yyy]
	
		#print("PCI:",p)
		temp.sort_values('Cell')
		temp.reset_index()
		#print(temp)
		#print(temp.iloc[0,0])

		#print('length:',len(temp['Cell']))
		#print(len(temp['Cell']))
		for i in range(len(temp['Cell'])-1):
	#source_cell.append(temp.iloc[i,0])
		#print('value of i:',i)
			for k in range(len(temp['Cell'])-(i+1)):
				source_cell.append((temp.iloc[i,0]))
				pci_final.append(p)
		#print("source cell", source_cell)
			#print('source cell',source_cell)
			for j in range(i+1,len(temp['Cell'])):
			#print("value of J",j)
				x=temp.iloc[j,0]
				
				target_cell.append(x)
			#print("target cell:",target_cell)
				lon_1=temp.iloc[i,2].astype(float)
				
				lat_1=temp.iloc[i,1].astype(float)
				lon_2=temp.iloc[j,2].astype(float)
				lat_2=temp.iloc[j,1].astype(float)
			#print("source coordinates",lat_1,lon_1)
			#print("target coordinates",lat_2,lon_2)
				d=haversine(lon_1,lat_1,lon_2,lat_2)
				distance.append(d)
			#print("distance b/w source and target",d)
			
	Min_dist=6
	df_f_1=pd.DataFrame(list(zip(source_cell,target_cell,pci_final,distance)), columns=['Source_Cell','Target_Cell','PCI','Distance(Km)']).sort_values(['Distance(Km)'])         
#print(df_f_1)

	data=df_f_1.groupby('PCI')['Distance(Km)'].agg(['min']).reset_index()


	df_f_2=tt.merge(data, on='PCI',how='inner')
	df_f_4=df_f_1[df_f_1['Distance(Km)']<Min_dist]
	df_f_5=df_f_4.groupby('PCI')['PCI'].agg(['count']).reset_index()
	df_f_5.columns=['PCI','Countofcells<DesignMin.Reuse']
	#df_f_2=pd.DataFrame(data)
	df_f_2.columns=['PCI','PCI_Reuse_Count','Min._Reuse_distance(km)']
	df_f_8=df_f_2.copy()
	del df_f_8['Min._Reuse_distance(km)']
	df_f_8=df_f_8.sort_values(by='PCI_Reuse_Count', ascending=False)
	print(df_f_8)
	df_f_3=df_f_2[df_f_2['Min._Reuse_distance(km)']<Min_dist]
	#print(df_f_3)
	#print(df_f_5)
	df_f_6=pd.merge(left=df_f_3,right=df_f_5, on='PCI')
	df_f_6=df_f_6.sort_values(by='Min._Reuse_distance(km)')
	print(df_f_6)
#### EXCEL EXPORTS
	output = BytesIO()
	writer = pd.ExcelWriter(output, engine='xlsxwriter')
	
	df_f_1.to_excel(writer,sheet_name='PCI_Reuse_Cell_details', index=False)
	df_f_8.to_excel(writer,sheet_name='PCI_Reuse_Stats',startrow=1, startcol=0,index=False)
	df_f_6.to_excel(writer,sheet_name='PCI_Reuse_Stats',startrow=1,startcol=4,index=False)

	workbook=writer.book
	worksheet1=writer.sheets['PCI_Reuse_Cell_details']

	worksheet1.set_column(0,0,12)
	worksheet1.set_column(1,1,12)
	worksheet1.set_column(2,2,12)
	worksheet1.set_column(3,3,20)
	header_format_2 = workbook.add_format({
	'bold': True,
	'text_wrap': False,
	'valign': 'top',
	'align': 'center',
	'fg_color': '#FFFF00',
	'border': 1})
	for col_num, value in enumerate(df_f_1.columns.values):
		worksheet1.write(0, col_num, value, header_format_2)  

##### WORKSHEET 2 FORMATTING

	workbook=writer.book
	worksheet=writer.sheets['PCI_Reuse_Stats']

	header_format = workbook.add_format({
	'bold': True,
	'text_wrap': False,
	'valign': 'top',
	'fg_color': '#D7E4BC',
	'border': 1})
	header_format_1 = workbook.add_format({
	'bold': True,
	'text_wrap': False,
	'valign': 'top',
	'fg_color': '#00FF00',
	'border': 1})
	header_format_2 = workbook.add_format({
	'bold': True,
	'text_wrap': False,
	'valign': 'top',
	'align': 'center',
	'fg_color': '#FFFF00',
	'border': 1})
	row_format_1 = workbook.add_format({
	'bold': False,
	'text_wrap': False,
	'valign': 'top',
	'align': 'center',
	'fg_color': '#808080',
	'border': 1})
	header_format_3 = workbook.add_format({
	'bold': True,
	'text_wrap': False,
	'valign': 'top',
	'align': 'center',
	'fg_color': '#FFB90F',
	'border': 1})
	header_format_7 = workbook.add_format({
	'bold': True,
	'text_wrap': False,
	'valign': 'top',
	'align': 'center',
	'fg_color': '#800000',
	'border': 1})   
	
	for col_num, value in enumerate(df_f_8.columns.values):
		worksheet.write(1, col_num+1, value, header_format)    
	for col_num,value in enumerate(df_f_6.columns.values):
		worksheet.write(1,col_num+4,value,header_format_1)
	worksheet.set_column(0,0,5)
	worksheet.set_column(1,1,16)
	worksheet.set_column(2,2,23)
	worksheet.set_column(4,4,5)
	worksheet.set_column(5,5,16)
	worksheet.set_column(6,6,23)
	worksheet.set_column(7,7,27)
	worksheet.set_column(3,3,3)
	worksheet.set_column(8,8,3)
	bold   = workbook.add_format({'bold': True})
	italic = workbook.add_format({'italic': True})
	worksheet.merge_range('E1:H1','PCIs with Reuse distance < Kms',header_format_2)

	worksheet.merge_range('J1:N1','Current PCI PLAN statistics',header_format_3)
	worksheet.merge_range('B1:C1',' Most Reused PCIs', header_format_2)
	worksheet.merge_range('J2:M2','Minimum Re-use distance achieved:',header_format_1)
	worksheet.merge_range('J3:M3','No. of cells with PCI reuse <5Kms',header_format_1)
	print(df_f_6)
	mini_reuse='%0.2f' %df_f_6['Min._Reuse_distance(km)'].head(1)
	print(mini_reuse)
	no_of_cells=df_f_6['Countofcells<DesignMin.Reuse'].sum()
	worksheet.write("N2",mini_reuse,header_format_2)
	worksheet.write("N3",no_of_cells,header_format_2)
	#worksheet.merge_range('D1:D100',header_format_7)
	#worksheet.merge_range('I1:I100',header_format_7)
	print(mini_reuse)
	# CHARTS CREATION AND FORMATTING


	chart=workbook.add_chart({'type':'bar'})
	chart.add_series({'values': ['PCI_Reuse_Stats',3,5,8,5], 'Categories': ['PCI_Reuse_Stats',3,4,8,4] })

	worksheet.insert_chart('J5', chart)
	writer.save()
	writer.close()
	workbook.close()
	
	output.seek(0)  
	#return Response(df1,headers={"Content-disposition":"attachment; filename=myplot.xls"}) 
	return send_file(output, attachment_filename="Re-use_details.xlsx", as_attachment=True)
	



@app.route('/audit_ca_lb_tdd', methods=['GET', 'POST'])

def audit_ca_lb_tdd():

	ca_lb_export_tdd=request.files['new_site_data_2']
	enbid_data=request.files['new_site_data_3']
	df_999=pd.read_excel(ca_lb_export_tdd)
	df_998=pd.read_excel(enbid_data)

	## Convert to string all major columns
	df_999['EUtranCellRelation']=df_999['EUtranCellRelation'].astype(str)
	df_999['EUtranFreqRelation']=df_999['EUtranFreqRelation'].astype(str)
	df_999['lbBnrAllowed']=df_999['lbBnrAllowed'].astype(str)
	df_999['EUtranCellTDD']=df_999['EUtranCellTDD'].astype(str)

	#DROP INTRA FREQ RELATIONS

	df_999=df_999.drop(df_999[np.logical_and(df_999['EUtranCellTDD'].str[-1:].str.contains('|'.join(['A','B','C'])), df_999['EUtranFreqRelation']=='1')].index)
	df_999=df_999.drop(df_999[np.logical_and(df_999['EUtranCellTDD'].str[-1:].str.contains('|'.join(['A','B','C'])), df_999['EUtranFreqRelation']=='39210')].index)
	df_999=df_999.drop(df_999[np.logical_and(df_999['EUtranCellTDD'].str[-1:].str.contains('|'.join(['E','F','G'])), df_999['EUtranFreqRelation']=='39408')].index)


	## Convert to string all major columns

	df_999.insert(4,'source_enbid',df_999['EUtranCellTDD'].map(df_998.set_index('cell')['enbid']))
	df_999['cellid_n']=df_999['EUtranCellRelation'].str[-1:].astype(str)
	df_999['source_enbid']=df_999['source_enbid'].astype(str)


	def col_check_tdd(xx):
		
		temp_enbid=xx['source_enbid']
	
		if ((temp_enbid in xx['EUtranCellRelation']) and (xx['EUtranFreqRelation']=='1')):
			xx['colocated']='colocated_L2300_F1'        
		elif ((temp_enbid in xx['EUtranCellRelation']) and (xx['EUtranFreqRelation']=='39210')):
			xx['colocated']='colocated_L2300_F1'
		elif ((temp_enbid in xx['EUtranCellRelation']) and (xx['EUtranFreqRelation']=='39408')):
			xx['colocated']='colocated_L2300_F2'
		else:
			xx['colocated']='Non-colocated' 
		return xx

	def lb_check_tdd(yy):
	
		t_enbid=yy['source_enbid']
		if(t_enbid in yy['EUtranCellRelation']):
	   
			if ((yy['EUtranCellTDD'][-1:]=='A') & (yy['cellid_n'] in (['3']))):
				yy['lb_colocation']='colocated_f2_sec1'
			elif ((yy['EUtranCellTDD'][-1:]=='B') & (yy['cellid_n'] in (['4']))):
				yy['lb_colocation']='colocated_f2_sec2'
			elif ((yy['EUtranCellTDD'][-1:]=='C') & (yy['cellid_n'] in (['5']))):
				yy['lb_colocation']='colocated_f2_sec3'
			elif ((yy['EUtranCellTDD'][-1:]=='E') & (yy['cellid_n'] in (['0']))):
				yy['lb_colocation']='colocated_f1_sec1'
			elif ((yy['EUtranCellTDD'][-1:]=='F') & (yy['cellid_n'] in (['1']))):
				yy['lb_colocation']='colocated_f1_sec2'
			elif ((yy['EUtranCellTDD'][-1:]=='G') & (yy['cellid_n'] in (['2']))):
				yy['lb_colocation']='colocated_f1_sec3'
			else:
				yy['lb_colocation']='Non-colocated' 
		else:
			yy['lb_colocation']='Non-colocated'
	
		return yy
	# create a new column of colocated or non-colocated
  
	df_333=df_999.apply(col_check_tdd,axis=1)
	df_333=df_333.apply(lb_check_tdd,axis=1)
			   
			   # Non-colocated Data
	df_444=df_333[df_333['colocated']=='Non-colocated']

			## Non-colocated Check
	## Non-colocated data where scellcandidate is enabled wrongly

	df_555=df_444[df_444['sCellCandidate']==1]
	noncol_wrong_scell_tdd=len(df_555['sCellCandidate'])
	df_555_new=df_555[['EUtranCellTDD','EUtranFreqRelation','EUtranCellRelation','sCellCandidate','source_enbid','colocated']]

			## COLOCATED DATA

	## Colocated data after removing Non-located relations

	df_666=df_333[df_333['colocated']!='Non-colocated']
	df_666['site']=df_666['EUtranCellTDD'].str[:6]

		  ### COLOCATED CHECK
		  
	## colocated data where Scellcandidate is NOT ENABLED

	df_777=df_666[df_666['sCellCandidate']!=1]
	df_777_new=df_777[['EUtranCellTDD','EUtranFreqRelation','EUtranCellRelation','sCellCandidate','source_enbid','colocated']]
	col_wrong_scell_tdd=len(df_777['sCellCandidate'])

	## colocated data where LB is wrongly set

	df_wrong_lb_tdd_temp=df_333[df_333['lb_colocation']!='Non-colocated']
	df_wrong_lb_tdd=df_wrong_lb_tdd_temp[df_wrong_lb_tdd_temp['loadBalancing']!=1]
	df_wrong_lb_tdd_new=df_wrong_lb_tdd[['EUtranCellTDD','EUtranFreqRelation','EUtranCellRelation','loadBalancing','cellid_n','lb_colocation']]
	col_wrong_lb_tdd=len(df_wrong_lb_tdd['lb_colocation'])

	## colocated data where LBBNR is TRUE

	df_wrong_bnr_tdd=df_wrong_lb_tdd_temp[df_wrong_lb_tdd_temp['lbBnrAllowed']=='True']
	df_wrong_bnr_tdd_new=df_wrong_bnr_tdd[['EUtranCellTDD','EUtranFreqRelation','EUtranCellRelation','loadBalancing','lbBnrAllowed','cellid_n','lb_colocation']]
	col_wrong_bnr_tdd=len(df_wrong_bnr_tdd['lbBnrAllowed'])

	### colocated data where Scellprio is wrong

	df_wrong_scellprio_tdd=df_wrong_lb_tdd_temp[df_wrong_lb_tdd_temp['sCellPriority']!=7]
	df_wrong_scellprio_tdd_new=df_wrong_scellprio_tdd[['EUtranCellTDD','EUtranFreqRelation','EUtranCellRelation','sCellCandidate','sCellPriority','source_enbid','lb_colocation']]
	col_wrong_scellprio_tdd=len(df_wrong_scellprio_tdd['sCellPriority'])

	## Unique cells count and their actual colocated relations count
	df_888=df_666.groupby('EUtranCellTDD')['colocated'].count().reset_index()
	df_99=pd.DataFrame(df_888)
	df_99.columns=['cell','actual_relations']

	#co-sector missing relations
	df_col_rel=df_wrong_lb_tdd_temp.groupby('EUtranCellTDD')['lb_colocation'].count().reset_index()
	df_col_rel_missing=df_col_rel[df_col_rel['lb_colocation']==0]
	df_col_rel_missing['lb_colocation']=df_col_rel_missing['lb_colocation'].astype(str)
	df_col_rel_missing.columns=['Cell','Missing co-sector relations']


	## Calculating count per site
	uniq_cell_tdd=df_333['EUtranCellTDD'].unique()
	df_100=pd.DataFrame(uniq_cell_tdd)
	df_100['site']=df_100.iloc[:,0].str[:6]
	df_100.columns=['EUtranCellTDD','site']
	exp_uniq=df_100.groupby('site')['EUtranCellTDD'].size().reset_index()
	exp_uniq.columns=['site','count']


	## Calculating unqiue cells and their total expected colocated relations count
	act_uniq=df_333['EUtranCellTDD'].unique()
	act_df=pd.DataFrame(act_uniq)
	act_df.columns=['cell']
	act_df['site']=act_df['cell'].str[:6]
	act_df_new=act_df
	act_df_new.insert(2,'expected_relations', act_df['site'].map(exp_uniq.set_index('site')['count']))

	# subtracting by 3 to have expected relations based on technology
	act_df_new['expected_relations']=(np.subtract(act_df_new['expected_relations'],3)).astype(int)

	## single dataframe containing expected and actual relations
	act_df_new.insert(3,'actual_relations',act_df_new['cell'].map(df_99.set_index('cell')['actual_relations']))

	## fill NA values to 0
	act_df_new=act_df_new.fillna(0,axis=0)
	## calculating the np. of missing relations
	act_df_new['no_of_missing_relations']=act_df_new['expected_relations']-act_df_new['actual_relations']

	## missing colocated Relations
	miss_df_tdd=act_df_new[act_df_new['no_of_missing_relations']>0]
	no_missing_col_rel_tdd=len(miss_df_tdd['no_of_missing_relations'])

	## blank df for summary
	df_766=pd.DataFrame()
	output = BytesIO()
	writer = pd.ExcelWriter(output, engine='xlsxwriter')

	
	df_777_new.to_excel(writer,sheet_name='Col_wrong_Scell', index=False)
	df_555_new.to_excel(writer,sheet_name='Non-col_wrong_Scell',startrow=0, startcol=0,index=False)
	df_wrong_lb_tdd_new.to_excel(writer,sheet_name='Col_wrong_LB',startrow=0, startcol=0,index=False)
	df_wrong_bnr_tdd_new.to_excel(writer,sheet_name='Col_wrong_LBNR',startrow=0, startcol=0,index=False)
	df_wrong_scellprio_tdd_new.to_excel(writer,sheet_name='Col_wrong_ScellPrio',startrow=0, startcol=0,index=False)
	miss_df_tdd.to_excel(writer,sheet_name='Missing_colocated_relations',startrow=0,startcol=0,index=False)
	df_333.to_excel(writer,sheet_name='Raw_data',startrow=0,startcol=0,index=False)
	df_766.to_excel(writer,sheet_name='Summary_Statistics',startrow=0,startcol=0,index=False)
	df_col_rel_missing.to_excel(writer,sheet_name='Missing_colocated_relations',startrow=1,startcol=6,index=False)


	## formatting

	workbook=writer.book
	worksheet1=writer.sheets['Col_wrong_Scell']
	worksheet2=writer.sheets['Non-col_wrong_Scell']
	worksheet3=writer.sheets['Col_wrong_LB']
	worksheet4=writer.sheets['Col_wrong_LBNR']
	worksheet5=writer.sheets['Col_wrong_ScellPrio']
	worksheet6=writer.sheets['Missing_colocated_relations']
	worksheet7=writer.sheets['Raw_data']
	worksheet8=writer.sheets['Summary_Statistics']

	worksheet1.set_column(0,0,16)
	worksheet1.set_column(1,1,20)
	worksheet1.set_column(2,2,18)
	worksheet1.set_column(3,3,16)
	worksheet1.set_column(4,4,20)
	worksheet1.set_column(5,5,20)
	worksheet1.set_column(6,6,12)
	worksheet1.set_column(7,7,16)
	worksheet1.set_column(8,8,20)
	worksheet1.set_column(9,9,18)
	worksheet1.set_column(10,10,28)
	worksheet1.set_column(11,11,20)
	worksheet1.set_column(12,12,20)
	worksheet1.set_column(13,13,12)

	worksheet2.set_column(0,0,16)
	worksheet2.set_column(1,1,20)
	worksheet2.set_column(2,2,18)
	worksheet2.set_column(3,3,16)
	worksheet2.set_column(4,4,18)
	worksheet2.set_column(5,5,20)
	worksheet2.set_column(6,6,16)
	worksheet2.set_column(7,7,20)
	worksheet2.set_column(8,8,14)
	worksheet2.set_column(9,9,26)
	worksheet2.set_column(10,10,18)


	worksheet3.set_column(0,0,16)
	worksheet3.set_column(1,1,20)
	worksheet3.set_column(2,2,20)
	worksheet3.set_column(3,3,18)
	worksheet3.set_column(4,4,26)
	worksheet3.set_column(5,5,20)
	worksheet3.set_column(6,6,16)	
	worksheet3.set_column(7,7,20)
	worksheet3.set_column(8,8,14)
	worksheet3.set_column(9,9,26)
	worksheet3.set_column(10,10,18)

	worksheet4.set_column(0,0,16)
	worksheet4.set_column(1,1,20)
	worksheet4.set_column(2,2,20)
	worksheet4.set_column(3,3,18)
	worksheet4.set_column(4,4,18)
	worksheet4.set_column(5,5,20)
	worksheet4.set_column(6,6,16)
	worksheet4.set_column(7,7,20)
	worksheet4.set_column(8,8,14)
	worksheet4.set_column(9,9,26)
	worksheet4.set_column(10,10,18)
	worksheet4.set_column(11,11,18)

	worksheet5.set_column(0,0,16)
	worksheet5.set_column(1,1,20)
	worksheet5.set_column(2,2,20)
	worksheet5.set_column(3,3,18)
	worksheet5.set_column(4,4,18)
	worksheet5.set_column(5,5,20)
	worksheet5.set_column(6,6,16)
	worksheet5.set_column(7,7,20)
	worksheet5.set_column(8,8,14)
	worksheet5.set_column(9,9,26)
	worksheet5.set_column(10,10,18)
	worksheet5.set_column(11,11,18)

	worksheet6.set_column(0,0,16)
	worksheet6.set_column(1,1,20)
	worksheet6.set_column(2,2,20)
	worksheet6.set_column(3,3,18)
	worksheet6.set_column(4,4,18)
	worksheet6.set_column(5,5,20)
	worksheet6.set_column(6,6,16)
	worksheet6.set_column(7,7,20)
	worksheet6.set_column(8,8,14)
	worksheet6.set_column(9,9,26)
	worksheet6.set_column(10,10,18)


	header_format = workbook.add_format({
		'bold': True,
		'text_wrap': False,
		'valign': 'top',
		'fg_color': '#D7E4BC',
		'border': 1})
	header_format_1 = workbook.add_format({
		'bold': True,
		'text_wrap': False,
		'align': 'center',
		'valign': 'top',
		'fg_color': '#00FF00',
		'border': 1})
	header_format_2 = workbook.add_format({
		'bold': True,
		'text_wrap': False,
		'valign': 'top',
		'align': 'center',
		'fg_color': '#FFFF00',
		'border': 1})
	row_format_1 = workbook.add_format({
		'bold': False,
		'text_wrap': False,
		'valign': 'top',
		'align': 'center',
		'fg_color': '#808080',
		'border': 1})
	header_format_3 = workbook.add_format({
		'bold': True,
		'text_wrap': False,
		'valign': 'top',
		'align': 'center',
		'fg_color': '#FFB90F',
		'border': 1})
	header_format_7 = workbook.add_format({
		'bold': True,
		'text_wrap': False,
		'valign': 'center',
		'align': 'center',
		'fg_color': '#9ebaf6',
		'border': 1})   
	
	for col_num, value in enumerate(df_777_new.columns.values):
		worksheet1.write(0, col_num, value, header_format_1)    
	for col_num,value in enumerate(df_555_new.columns.values):
		worksheet2.write(0,col_num,value,header_format_1)
	for col_num, value in enumerate(df_wrong_lb_tdd_new.columns.values):
		worksheet3.write(0, col_num, value, header_format_1)    
	for col_num,value in enumerate(df_wrong_bnr_tdd_new.columns.values):
		worksheet4.write(0,col_num,value,header_format_1)
	for col_num,value in enumerate(df_wrong_scellprio_tdd_new.columns.values):
		worksheet5.write(0,col_num,value,header_format_1)
	for col_num, value in enumerate(miss_df_tdd.columns.values):
		worksheet6.write(0, col_num, value, header_format_1)    
	for col_num,value in enumerate(df_333.columns.values):
		worksheet7.write(0,col_num,value,header_format_1)
	for col_num,value in enumerate(df_766.columns.values):
		worksheet8.write(0,col_num,value,header_format_1)   
	for col_num,value in enumerate(df_col_rel_missing.columns.values):
		worksheet6.write(1,col_num+6,value,header_format_1)   
	
	worksheet6.merge_range('G1:H1','Total no. of missing co-secor relations',header_format_1)  


	worksheet8.merge_range('J2:Q2','SUMMARY STATISTICS',header_format_1)      
	worksheet8.merge_range('J3:O3','Collocated cell relations with wrong Scell value',header_format_2)
	worksheet8.merge_range('J4:O4','Non-collocated cell relations having wrong Scell value',header_format_2)
	worksheet8.merge_range('J8:O8','Missing collocated cell relations(Not defined)',header_format_2)
	worksheet8.merge_range('J6:O6','Collocated cell relations with wrong Load balancing',header_format_2)
	worksheet8.merge_range('J7:O7','Collocated cell relations with wrong LBNR value',header_format_2)
	worksheet8.merge_range('J5:O5','Collocated cell relations with wrong Scellprio value',header_format_2)

	worksheet8.merge_range('P3:Q3',col_wrong_scell_tdd,header_format_3)
	worksheet8.merge_range('P4:Q4',noncol_wrong_scell_tdd,header_format_3)
	worksheet8.merge_range('P5:Q5',col_wrong_scellprio_tdd,header_format_3)
	worksheet8.merge_range('P6:Q6',col_wrong_lb_tdd,header_format_3)
	worksheet8.merge_range('P7:Q7',col_wrong_bnr_tdd,header_format_3)
	worksheet8.merge_range('P8:Q8',col_wrong_scellprio_tdd,header_format_3)
	worksheet8.merge_range('P8:Q8',no_missing_col_rel_tdd,header_format_3)

	worksheet8.merge_range('G3:I5','Carrier-Aggregation',header_format_7)
	worksheet8.merge_range('G6:I7','Load-Balancing',header_format_7)
	worksheet8.merge_range('G8:I8','Missing-Relations',header_format_7)

	writer.save()
	writer.close()
	workbook.close()
	output.seek(0)  
	
	return send_file(output, attachment_filename="CA_LB_Audit_Report_TDD.xlsx", as_attachment=True)

@app.route('/audit_ca_lb_fdd', methods=['GET', 'POST'])

def audit_ca_lb_fdd():

	ca_lb_export_fdd=request.files['new_site_data_4']
	enbid_data=request.files['new_site_data_5']
	df_1=pd.read_excel(ca_lb_export_fdd)
	df_2=pd.read_excel(enbid_data)


	## Convert to string all major columns

	df_1.insert(4,'source_enbid',df_1['EUtranCellFDD'].map(df_2.set_index('cell')['enbid']))
	df_1['EUtranCellRelation']=df_1['EUtranCellRelation'].astype(str)
	df_1['EUtranFreqRelation']=df_1['EUtranFreqRelation'].astype(str)
	df_1['source_enbid']=df_1['source_enbid'].astype(str)
	df_1['lbBnrAllowed']=df_1['lbBnrAllowed'].astype(str)
	df_1['cellid_n']=df_1['EUtranCellRelation'].str[-1:].astype(str)
	df_1['EUtranCellFDD']=df_1['EUtranCellFDD'].astype(str)
	def col_check(xx):
	
		temp_enbid=xx['source_enbid']
	
		if ((temp_enbid in xx['EUtranCellRelation']) and (xx['EUtranFreqRelation']=='1450')):
			xx['colocated']='colocated_L1800'        
		elif ((temp_enbid in xx['EUtranCellRelation']) and (xx['EUtranFreqRelation']=='451')):
			xx['colocated']='colocated_L2100'
		elif ((temp_enbid in xx['EUtranCellRelation']) and (xx['EUtranFreqRelation']=='9310')):
			xx['colocated']='colocated_L700'
		else:
			xx['colocated']='Non-colocated' 

		return xx

	def lb_check(yy):
	
		t_enbid=yy['source_enbid']
		if(t_enbid in yy['EUtranCellRelation']):
	   
			if ((yy['EUtranCellFDD'][-1:]=='A') & (yy['cellid_n'] in (['0','3','6']))):
				yy['lb_colocation']='colocated_sec1'
			elif ((yy['EUtranCellFDD'][-1:]=='B') & (yy['cellid_n'] in (['1','4','7']))):
				yy['lb_colocation']='colocated_sec2'
			elif ((yy['EUtranCellFDD'][-1:]=='C') & (yy['cellid_n'] in (['2','5','8']))):
				yy['lb_colocation']='colocated_sec3'
			else:
				yy['lb_colocation']='Non-colocated' 
		else:
			yy['lb_colocation']='Non-colocated'
	
		return yy

	# create a new column of colocated or non-colocated
  
	df_3=df_1.apply(col_check,axis=1)
	
	df_3=df_3.apply(lb_check,axis=1)
			
			  # Non-colocated DataFRAME

	df_4=df_3[df_3['colocated']=='Non-colocated']

			## NON COLOCATED CHECK ####

	## Non-colocated data where scellcandidate is enabled wrongly

	df_5=df_4[df_4['sCellCandidate']==1]
	df_5_new=df_5[['EUtranCellFDD','EUtranFreqRelation','EUtranCellRelation','sCellCandidate','source_enbid','colocated']]
	noncol_wrong_scell_fdd=len(df_5['sCellCandidate'])

				### COLOCATED DATAFRAME
	## Colocated data after removing Non-located relations

	df_6=df_3[df_3['colocated']!='Non-colocated']
	df_6['site']=df_6['EUtranCellFDD'].str[:6]


				 ## COLOCATED CHECK ####
	## colocated data where Scellcandidate is NOT ENABLED

	df_7=df_6[df_6['sCellCandidate']!=1]
	df_7_new=df_7[['EUtranCellFDD','EUtranFreqRelation','EUtranCellRelation','sCellCandidate','source_enbid','colocated']]
	col_wrong_scell_fdd=len(df_7['sCellCandidate'])

	## colocated data where loadbalancing is not enabled

	df_wrong_lb_fdd_temp=df_3[df_3['lb_colocation']!='Non-colocated']
	df_wrong_lb_fdd=df_wrong_lb_fdd_temp[df_wrong_lb_fdd_temp['loadBalancing']!=1]
	df_wrong_lb_fdd_new=df_wrong_lb_fdd[['EUtranCellFDD','EUtranFreqRelation','EUtranCellRelation','loadBalancing','cellid_n','lb_colocation']]
	col_wrong_lb_fdd=len(df_wrong_lb_fdd['lb_colocation'])

	## colocated data where LBBNR is TRUE

	df_wrong_bnr_fdd=df_wrong_lb_fdd_temp[df_wrong_lb_fdd_temp['lbBnrAllowed']=='True']
	df_wrong_bnr_fdd_new=df_wrong_bnr_fdd[['EUtranCellFDD','EUtranFreqRelation','EUtranCellRelation','loadBalancing','lbBnrAllowed','cellid_n','lb_colocation']]
	col_wrong_bnr_fdd=len(df_wrong_bnr_fdd['lbBnrAllowed'])

	## colocated data where Scellprio is not 7

	df_wrong_scellprio_fdd=df_wrong_lb_fdd_temp[df_wrong_lb_fdd_temp['sCellPriority']!=7]
	df_wrong_scellprio_fdd_new=df_wrong_scellprio_fdd[['EUtranCellFDD','EUtranFreqRelation','EUtranCellRelation','sCellCandidate','sCellPriority','source_enbid','lb_colocation']]
	col_wrong_scellprio_fdd=len(df_wrong_scellprio_fdd['sCellPriority'])

	## Unique cells count and their actual colocated relations count

	df_8=df_6.groupby('EUtranCellFDD')['colocated'].count().reset_index()
	df_9=pd.DataFrame(df_8)
	df_9.columns=['cell','actual_relations']

	## Calculating count per site

	uniq_cell=df_3['EUtranCellFDD'].unique()
	df_16=pd.DataFrame(uniq_cell)
	df_16['site']=df_16.iloc[:,0].str[:6]	
	df_16.columns=['cell','site']
	df_16['site']=df_16['site'].str.replace('F','E').str.replace('X','E')
	exp_uniq=df_16.groupby('site')['cell'].size().reset_index()

	## Calculating unqiue cells and their total expected colocated relations count
	
	act_uniq=df_3['EUtranCellFDD'].unique()	
	act_df=pd.DataFrame(act_uniq)
	act_df.columns=['cell']
	act_df['site']=act_df['cell'].str[:6].str.replace('F','E').str.replace('X','E')
	act_df.insert(2,'expected_relations', act_df['site'].map(exp_uniq.set_index('site')['cell']))

	# dividing by 3 to have expected relations based on technology

	act_df['expected_relations']=(np.subtract(act_df['expected_relations'],3)).astype(int)

	## single dataframe containing expected and actual relations
	
	act_df.insert(3,'actual_relations',act_df['cell'].map(df_9.set_index('cell')['actual_relations']))

	## filling NA values to 0
	act_df=act_df.fillna(0, axis=0)
	
	## calculating the np. of missing relations
	act_df['no_of_missing_relations']=act_df['expected_relations']-act_df['actual_relations']

	## missing colocated Relations
	miss_df=act_df[act_df['no_of_missing_relations']>0]
	no_missing_col_rel_fdd=len(miss_df['no_of_missing_relations'])

	## blank df
	df_788=pd.DataFrame()
	output = BytesIO()
	writer = pd.ExcelWriter(output, engine='xlsxwriter')

	
	df_7_new.to_excel(writer,sheet_name='Col_wrong_Scell', index=False)
	df_5_new.to_excel(writer,sheet_name='Non-col_wrong_Scell',startrow=0, startcol=0,index=False)
	df_wrong_lb_fdd_new.to_excel(writer,sheet_name='Col_wrong_LB',startrow=0, startcol=0,index=False)
	df_wrong_bnr_fdd_new.to_excel(writer,sheet_name='Col_wrong_LBNR',startrow=0, startcol=0,index=False)
	df_wrong_scellprio_fdd_new.to_excel(writer,sheet_name='Col_wrong_ScellPrio',startrow=0, startcol=0,index=False)
	miss_df.to_excel(writer,sheet_name='Missing_colocated_relations',startrow=0,startcol=0,index=False)
	df_3.to_excel(writer,sheet_name='Raw_data',startrow=0,startcol=0,index=False)
	df_788.to_excel(writer,sheet_name='Summary_Statistics',startrow=0,startcol=0,index=False)



	## formatting

	workbook=writer.book
	worksheet1=writer.sheets['Col_wrong_Scell']
	worksheet2=writer.sheets['Non-col_wrong_Scell']
	worksheet3=writer.sheets['Col_wrong_LB']
	worksheet4=writer.sheets['Col_wrong_LBNR']
	worksheet5=writer.sheets['Col_wrong_ScellPrio']
	worksheet6=writer.sheets['Missing_colocated_relations']
	worksheet7=writer.sheets['Raw_data']
	worksheet8=writer.sheets['Summary_Statistics']

	worksheet1.set_column(0,0,16)
	worksheet1.set_column(1,1,20)
	worksheet1.set_column(2,2,18)
	worksheet1.set_column(3,3,16)
	worksheet1.set_column(4,4,20)
	worksheet1.set_column(5,5,20)
	worksheet1.set_column(6,6,12)
	worksheet1.set_column(7,7,16)
	worksheet1.set_column(8,8,20)
	worksheet1.set_column(9,9,18)
	worksheet1.set_column(10,10,28)
	worksheet1.set_column(11,11,20)
	worksheet1.set_column(12,12,20)
	worksheet1.set_column(13,13,12)

	worksheet2.set_column(0,0,16)
	worksheet2.set_column(1,1,20)
	worksheet2.set_column(2,2,18)
	worksheet2.set_column(3,3,16)
	worksheet2.set_column(4,4,18)
	worksheet2.set_column(5,5,20)
	worksheet2.set_column(6,6,16)
	worksheet2.set_column(7,7,20)
	worksheet2.set_column(8,8,14)
	worksheet2.set_column(9,9,26)
	worksheet2.set_column(10,10,18)


	worksheet3.set_column(0,0,16)
	worksheet3.set_column(1,1,20)
	worksheet3.set_column(2,2,20)
	worksheet3.set_column(3,3,18)
	worksheet3.set_column(4,4,26)
	worksheet3.set_column(5,5,20)
	worksheet3.set_column(6,6,16)
	worksheet3.set_column(7,7,20)
	worksheet3.set_column(8,8,14)
	worksheet3.set_column(9,9,26)
	worksheet3.set_column(10,10,18)

	worksheet4.set_column(0,0,16)
	worksheet4.set_column(1,1,20)
	worksheet4.set_column(2,2,14)
	worksheet4.set_column(3,3,16)
	worksheet4.set_column(4,4,18)
	worksheet4.set_column(5,5,20)
	worksheet4.set_column(6,6,16)
	worksheet4.set_column(7,7,20)
	worksheet4.set_column(8,8,14)
	worksheet4.set_column(9,9,26)	
	worksheet4.set_column(10,10,18)
	worksheet4.set_column(11,11,18)

	worksheet5.set_column(0,0,16)
	worksheet5.set_column(1,1,20)
	worksheet5.set_column(2,2,14)
	worksheet5.set_column(3,3,16)
	worksheet5.set_column(4,4,18)
	worksheet5.set_column(5,5,20)
	worksheet5.set_column(6,6,16)	
	worksheet5.set_column(7,7,20)
	worksheet5.set_column(8,8,14)
	worksheet5.set_column(9,9,26)
	worksheet5.set_column(10,10,18)
	worksheet5.set_column(11,11,18)

	worksheet6.set_column(0,0,16)
	worksheet6.set_column(1,1,20)
	worksheet6.set_column(2,2,14)
	worksheet6.set_column(3,3,16)
	worksheet6.set_column(4,4,18)
	worksheet6.set_column(5,5,20)
	worksheet6.set_column(6,6,16)
	worksheet6.set_column(7,7,20)
	worksheet6.set_column(8,8,14)
	worksheet6.set_column(9,9,26)
	worksheet6.set_column(10,10,18)

	worksheet7.set_column(0,0,16)
	worksheet7.set_column(1,1,20)
	worksheet7.set_column(2,2,14)
	worksheet7.set_column(3,3,16)
	worksheet7.set_column(4,4,18)
	worksheet7.set_column(5,5,20)
	worksheet7.set_column(6,6,16)
	worksheet7.set_column(7,7,20)
	worksheet7.set_column(8,8,14)
	worksheet7.set_column(9,9,26)
	worksheet7.set_column(10,10,18)


	header_format = workbook.add_format({
		'bold': True,
		'text_wrap': False,
		'valign': 'top',
		'fg_color': '#D7E4BC',
		'border': 1})
	header_format_1 = workbook.add_format({
		'bold': True,
		'text_wrap': False,
		'align': 'center',
		'valign': 'top',
		'fg_color': '#00FF00',
		'border': 1})
	header_format_2 = workbook.add_format({
		'bold': True,
		'text_wrap': False,
		'valign': 'top',
		'align': 'center',
		'fg_color': '#FFFF00',
		'border': 1})
	row_format_1 = workbook.add_format({
		'bold': False,
		'text_wrap': False,
		'valign': 'top',
		'align': 'center',
		'fg_color': '#808080',
		'border': 1})
	header_format_3 = workbook.add_format({
		'bold': True,
		'text_wrap': False,
		'valign': 'top',
		'align': 'center',
		'fg_color': '#FFB90F',
		'border': 1})
	header_format_7 = workbook.add_format({
		'bold': True,
		'text_wrap': False,
		'valign': 'top',
		'align': 'center',
		'fg_color': '#9ebaf6',
		'border': 1})   
	
	for col_num, value in enumerate(df_7_new.columns.values):
		worksheet1.write(0, col_num, value, header_format_1)    
	for col_num,value in enumerate(df_5_new.columns.values):
		worksheet2.write(0,col_num,value,header_format_1)
	for col_num, value in enumerate(df_wrong_lb_fdd_new.columns.values):
		worksheet3.write(0, col_num, value, header_format_1)    
	for col_num,value in enumerate(df_wrong_bnr_fdd_new.columns.values):
		worksheet4.write(0,col_num,value,header_format_1)
	for col_num,value in enumerate(df_wrong_scellprio_fdd_new.columns.values):
		worksheet5.write(0,col_num,value,header_format_1)
	for col_num,value in enumerate(miss_df.columns.values):
		worksheet6.write(0,col_num,value,header_format_1)
	for col_num,value in enumerate(df_3.columns.values):
		worksheet7.write(0,col_num,value,header_format_1)
	for col_num,value in enumerate(df_788.columns.values):
		worksheet8.write(0,col_num,value,header_format_1)
					
   
  
	worksheet8.merge_range('J2:Q2','SUMMARY STATISTICS',header_format_1)      	
	worksheet8.merge_range('J3:O3','No. of collocated cell relations having wrong Scell value',header_format_2)
	worksheet8.merge_range('J4:O4','No. of Non-collocated cell relations having wrong Scell value',header_format_2)
	worksheet8.merge_range('J8:O8','No. of missing collocated cell relations(Not defined)',header_format_2)
	worksheet8.merge_range('J6:O6','No. of colocated cell relations having wrong Load balancing',header_format_2)
	worksheet8.merge_range('J7:O7','No. of colocated cell relations having wrong LBNR value',header_format_2)
	worksheet8.merge_range('J5:O5','No. of colocated cell relations having wrong Scellprio value',header_format_2)


	worksheet8.merge_range('P3:Q3',col_wrong_scell_fdd,header_format_3)
	worksheet8.merge_range('P4:Q4',noncol_wrong_scell_fdd,header_format_3)
	worksheet8.merge_range('P5:Q5',col_wrong_scellprio_fdd,header_format_3)
	worksheet8.merge_range('P6:Q6',col_wrong_lb_fdd,header_format_3)
	worksheet8.merge_range('P7:Q7',col_wrong_bnr_fdd,header_format_3)
	worksheet8.merge_range('P8:Q8',no_missing_col_rel_fdd,header_format_3)


	worksheet8.merge_range('G3:I5','Carrier-Aggregation',header_format_7)
	worksheet8.merge_range('G6:I7','Load-Balancing',header_format_7)
	worksheet8.merge_range('G8:I8','Missing-Relations',header_format_7)

	writer.save()
	writer.close()
	workbook.close()
	output.seek(0)  
	
	return send_file(output, attachment_filename="CA_LB_Audit_Report_FDD.xlsx", as_attachment=True)

if __name__=="__main__":
	app.run(debug=True)
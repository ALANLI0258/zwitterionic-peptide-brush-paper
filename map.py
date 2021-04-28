import numpy as np
#import os
import xlwings as xw
#import linecache



def DATALOAD(_dis, _salt):

    dict = {'A': 0, 'B': 0.0385, 'C': 0.0769, 'D': 0.115, 'E': 0.154, 'F': 0.192, 'G': 0.231, 'H': 0.269, 'I': 0.308, 'J': 0.346, 'K': 0.385, 'L': 0.423, 'M': 0.462, 'N': 0.5, 'O': 0.538, 'P': 0.577, 'Q': 0.615, 'R':0.654, 'S': 0.692, 'T': 0.731, 'U': 0.769, 'V':0.808, 'W': 0.846, 'X': 0.885, 'Y': 0.923, 'Z': 0.962, 'a': 1, 'b': 1.04, 'c': 1.08, 'd': 1.12, 'e': 1.15, 'f': 1.19, 'g': 1.23, 'h': 1.27, 'i': 1.31, 'j': 1.35, 'k': 1.38, 'l': 1.42, 'm': 1.46, 'n': 1.5, '\"': 'NA', '\'': 'NA', ',': 'NA'}

    _linenum = 66
    if _salt==0:
        sheet_name_1 = str(_dis)
        sheet_name_2 = str(_dis)+'(num)'
    else:
        sheet_name_1 = str(_dis)+';'+str(_salt)+'(NACL)'
        sheet_name_2 = str(_dis)+';'+str(_salt)+'NACL'+'(num)'        
    

    wb.sheets.add(sheet_name_1)
    wb.sheets.add(sheet_name_2)


    #sht = wb.sheets[sheet_name]
    lines = []
    lnum = 0
    with open(file_path+'map.xpm','r') as f:
        for line in f:
            lnum += 1
            if (lnum >_linenum):
                lines.append(line)
    list = lines

    titles1 = [[str(_salt)+'NaCl']]
    wb.sheets[sheet_name_1].range(1,1).value = titles1
    wb.sheets[sheet_name_2].range(1,1).value = titles1

    for i in range(len(list)):
        list_temp_1 = list[i].split()   
        list_temp_2 = list_temp_1[0]  
        list_temp_3 = []
        list_temp_4 = []
        #print (list_temp_2)
        for k in range(len(list_temp_2)):
            list_temp_3.append(list_temp_2[k])            
            #print (list_temp_3)
            temp=dict[list_temp_2[k]]
            list_temp_4.append(temp)
        wb.sheets[sheet_name_1].range(i+2,1).value = list_temp_3
        wb.sheets[sheet_name_2].range(i+2,1).value = list_temp_4

        #list[i]=list[i].split()
        #print(list[i])
        #wb.sheets[sheet_name].range(int(i+2),1).value = list[i]

    wb.sheets[sheet_name_2].api.rows('1').delete
    wb.sheets[sheet_name_2].api.columns('A').delete
    wb.sheets[sheet_name_1].range('A1:LP326').columns.autofit()
    wb.sheets[sheet_name_2].range('A1:LP326').columns.autofit()
    
    f.close()

def CHAIN_CHAIN(_dis,_interval,_sheet_name,_ychain,_xchain,_pairs_num):

    if _dis == 1.0:
        colunm_dis=1
        row_dis=32
        grafting_dens='1.00'
    elif _dis == 1.2:
        colunm_dis=8
        row_dis=33
        grafting_dens='0.69'
    elif _dis == 1.34:
        colunm_dis=15
        row_dis=34
        grafting_dens='0.56'
    elif _dis == 1.62:
        colunm_dis=22 
        row_dis=35   
        grafting_dens='0.38'
    elif _dis == 1.74:
        colunm_dis=29
        row_dis=36 
        grafting_dens='0.33'
    elif _dis == 1.8:
        colunm_dis=36
        row_dis=37  
        grafting_dens='0.31'       
    elif _dis == 2.0:
        colunm_dis=43
        row_dis=38
        grafting_dens='0.25'

    if _interval==1:
        colunm_interval=10
        row_interval=41
    elif _interval==1.414:
        colunm_interval=17
        row_interval=93
    elif _interval==2:
        colunm_interval=24
        row_interval=145

    head_tail_or_not_num=0
    parallel_or_not_num=0
    head_head_or_not_num=0

    head_tail_pairs_list=[]
    parallel_pairs_list=[]
    head_head_pairs_list=[]

    for l in range(len(_ychain)):
        head_tail_rightdown=[]
        head_tail_leftup=[]
        head_tail_rightdown_squar=[]
        head_tail_leftup_squar=[]        

        parallel_rightdown=[]
        parallel_leftup=[]
        parallel_middle=[]

        head_head_rightup=[]   
        head_head_leftup=[] 
        head_head_rightdown=[]   
        head_head_leftdown=[]
        head_head_down=[]

        #head_tail 
        #head_tail 
        #head_tail 
        i=0
        while i<12:
            row=(26-_ychain[l])*13-i
            j=12-i
            while j>0:
                column=_xchain[l]*13-j+1
                head_tail_rightdown.append(wb.sheets[_sheet_name].range(row,column).value)
                j-=1
            i+=1

        i=0
        while i<12:
            row=(25-_ychain[l])*13+i+1
            j=12-i
            while j>0:
                column=(_xchain[l]-1)*13+j
                head_tail_leftup.append(wb.sheets[_sheet_name].range(row,column).value)
                j-=1
            i+=1

        i=0
        while i<5:
            row=(25-_ychain[l])*13+i+1
            j=0
            while j<5:
                column=(_xchain[l]-1)*13+j+1
                head_tail_leftup_squar.append(wb.sheets[_sheet_name].range(row,column).value)
                j+=1
            i+=1

        i=0
        while i<5:
            row=(26-_ychain[l])*13-i
            j=0
            while j<5:
                column=_xchain[l]*13-j
                head_tail_rightdown_squar.append(wb.sheets[_sheet_name].range(row,column).value)
                j+=1
            i+=1

        head_tail_rightdown_dens=np.sum(list(map(lambda x: x < 1, head_tail_rightdown)))
        head_tail_leftup_dens=np.sum(list(map(lambda x: x < 1, head_tail_leftup)))
        head_tail_rightdown_squar_dens=np.sum(list(map(lambda x: x < 1, head_tail_rightdown_squar)))
        head_tail_leftup_squar_dens=np.sum(list(map(lambda x: x < 1, head_tail_leftup_squar)))        


        #print('head_tail_rightdown_dens')
        #print(head_tail_rightdown_dens)
        #print('head_tail_leftup_dens')
        #print(head_tail_leftup_dens)


        if (head_tail_rightdown_squar_dens>11 and head_tail_leftup_dens<15) or (head_tail_rightdown_dens<15 and head_tail_leftup_squar_dens>11):
            head_tail_or_not_num+=1
            head_tail_pairs=str(_ychain[l])+','+str(_xchain[l])
            head_tail_pairs_list.append(head_tail_pairs)
        else:
            pass

        #parallel
        #parallel
        #parallel
        i=0
        while i<9:
            row=(26-_ychain[l])*13-i
            j=9-i
            while j>0:
                column=_xchain[l]*13-j+1
                parallel_rightdown.append(wb.sheets[_sheet_name].range(row,column).value)
                j-=1
            i+=1

        i=0
        while i<9:
            row=(25-_ychain[l])*13+i+1
            j=9-i
            while j>0:
                column=(_xchain[l]-1)*13+j
                parallel_leftup.append(wb.sheets[_sheet_name].range(row,column).value)
                j-=1
            i+=1

        i=0
        while i<13:
            row=(26-_ychain[l])*13-i
            if i==0:
                column=(_xchain[l]-1)*13+1
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column).value)
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column+1).value)
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column+2).value)
            elif i==1:
                column=(_xchain[l]-1)*13+1
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column).value)
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column+1).value)
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column+2).value)  
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column+3).value)  
            elif i==11:
                column=(_xchain[l]-1)*13+10
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column).value)
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column+1).value)
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column+2).value)  
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column+3).value) 
            elif i==12:
                column=(_xchain[l]-1)*13+11
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column).value)
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column+1).value)
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column+2).value)  
            else:
                column=(_xchain[l]-1)*13+1+i
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column-2).value)
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column-1).value)
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column).value)
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column+1).value)
                parallel_middle.append(wb.sheets[_sheet_name].range(row,column+2).value) 

            i+=1

        parallel_rightdown_dens=np.sum(list(map(lambda x: x < 1, parallel_rightdown)))
        parallel_leftup_dens=np.sum(list(map(lambda x: x < 1, parallel_leftup)))
        parallel_middle_dens=np.sum(list(map(lambda x: x < 1, parallel_middle)))

        #print('parallel_rightdown_dens')
        #print(parallel_rightdown_dens)
        #print('parallel_leftup_dens')
        #print(parallel_leftup_dens)        
        #print('parallel_middle_dens')
        #print(parallel_middle_dens)  
        

        if parallel_rightdown_dens<10 and parallel_leftup_dens<10 and parallel_middle_dens>25:
            parallel_or_not_num+=1
            parallel_pairs=str(_ychain[l])+','+str(_xchain[l])
            parallel_pairs_list.append(parallel_pairs)


        #headhead
        #headhead
        #headhead

        i=0
        while i<5:
            row=(25-_ychain[l])*13+i+1
            j=0
            while j<5:
                column=_xchain[l]*13-j
                head_head_rightup.append(wb.sheets[_sheet_name].range(row,column).value)
                j+=1
            i+=1

        i=0
        while i<5:
            row=(25-_ychain[l])*13+i+1
            j=0
            while j<5:
                column=(_xchain[l]-1)*13+j+1
                head_head_leftup.append(wb.sheets[_sheet_name].range(row,column).value)
                j+=1
            i+=1

        i=0
        while i<8:
            row=(26-_ychain[l])*13-i
            j=0
            while j<13:
                column=(_xchain[l]-1)*13+j+1
                head_head_down.append(wb.sheets[_sheet_name].range(row,column).value)
                j+=1
            i+=1
                

        head_head_rightup_dens=np.sum(list(map(lambda x: x < 1, head_head_rightup)))
        head_head_leftup_dens=np.sum(list(map(lambda x: x < 1, head_head_leftup)))
        head_head_down_dens=np.sum(list(map(lambda x: x < 1, head_head_down)))

        if head_head_down_dens<9 and head_head_leftup_dens<10 and head_head_rightup_dens>8:
            head_head_or_not_num+=1
            head_head_pairs=str(_ychain[l])+','+str(_xchain[l])
            head_head_pairs_list.append(head_head_pairs)

        #print('head_head_rightdown_dens')
        #print(head_head_rightdown_dens)
        #print('head_head_leftdown_dens')
        #print(head_head_leftdown_dens)        
        #print('head_head_rightup_dens')
        #print(head_head_rightup_dens) 
        #print('head_head_leftup_dens')
        #print(head_head_leftup_dens)       
        #print('head_head_down_dens') 
        #print(head_head_down_dens)   

        print('\npairs'+str(_interval))
        print(str(_ychain[l])+','+str(_xchain[l]))
        print('head_tail_num') 
        print(head_tail_or_not_num)  
        print('parallel_num') 
        print(parallel_or_not_num)  
        print('head_head_num') 
        print(head_head_or_not_num)  


    head_tail_or_not_per=head_tail_or_not_num/_pairs_num
    parallel_or_not_per=parallel_or_not_num/_pairs_num
    head_head_or_not_per=head_head_or_not_num/_pairs_num

    wb.sheets['data'].range(row_dis,colunm_interval).value=[[head_tail_or_not_num, parallel_or_not_num, head_head_or_not_num, head_tail_or_not_per, parallel_or_not_per, head_head_or_not_per]]


    wb.sheets['data'].range(row_interval,colunm_dis).value=str(_dis)+'nm('+str(_interval)+')'   

    wb.sheets['data'].range(row_interval+1,colunm_dis).value=[['head-tail pairs (y,x)', 'parallel pairs (y,x)', 'head-head pairs (y,x)']] 



    if np.size(head_tail_pairs_list)==0:
        wb.sheets['data'].range(row_interval+2,colunm_dis).options(transpose=True).value=[['N.A.']]
    else:
        wb.sheets['data'].range(row_interval+2,colunm_dis).options(transpose=True).value=head_tail_pairs_list

    if np.size(parallel_pairs_list)==0:
        wb.sheets['data'].range(row_interval+2,colunm_dis+1).options(transpose=True).value=[['N.A.']]
    else:
        wb.sheets['data'].range(row_interval+2,colunm_dis+1).options(transpose=True).value=parallel_pairs_list   

    if np.size(head_head_pairs_list)==0:    
        wb.sheets['data'].range(row_interval+2,colunm_dis+2).options(transpose=True).value=[['N.A.']]
    else:
        wb.sheets['data'].range(row_interval+2,colunm_dis+2).options(transpose=True).value=head_head_pairs_list





def ANALYZE(_dis, _salt):

    pairs_y = []        
    pairs_x = []        
    pairs_num=0         

    pairs_nd_y = []       
    pairs_nd_x = []        
    pairs_nd_num=0         

    pairs_rd_y = []        
    pairs_rd_x = []        
    pairs_rd_num=0         

    if _salt==0:
        sheet_name = str(_dis)+'(num)'
    else:
        sheet_name = str(_dis)+';'+str(_salt)+'NACL'+'(num)'  

    if _dis == 1.0:
        colunm_dis=1
        row_dis=32
        grafting_dens='1.00'
    elif _dis == 1.2:
        colunm_dis=8
        row_dis=33
        grafting_dens='0.69'
    elif _dis == 1.34:
        colunm_dis=15
        row_dis=34
        grafting_dens='0.56'
    elif _dis == 1.62:
        colunm_dis=22 
        row_dis=35   
        grafting_dens='0.38'
    elif _dis == 1.74:
        colunm_dis=29
        row_dis=36 
        grafting_dens='0.33'
    elif _dis == 1.8:
        colunm_dis=36
        row_dis=37  
        grafting_dens='0.31'       
    elif _dis == 2.0:
        colunm_dis=43
        row_dis=38
        grafting_dens='0.25'


    titles1=str(_dis)+'nm'
    wb.sheets['data'].range(1,colunm_dis).value=titles1  

    titles2 = [['Chain No.', 'Loop', 'Loop Num', 'Loop Strength', 'Straight', 'Mediate']]
    wb.sheets['data'].range(2,colunm_dis).value=titles2

    titles3 = [['grafting density (nm-2)', 'distance (nm)', 'Loop', 'Straight', 'Mediate', 'Loop Rate', 'Straight Rate', 'Mediate Rate', '', 'Head-Tail(1)', 'Parallel(1)', 'Head-Head(1)', 'Head-Tail Rate(1)', 'Parallel Rate(1)', 'Head-Head Rate(1)', '', 'Head-Tail(root2)', 'Parallel(root2)', 'Head-Head(root2)', 'Head-Tail Rate(root2)', 'Parallel Rate(root2)', 'Head-Head Rate(root2)', '', 'Head-Tail(2)', 'Parallel(2)', 'Head-Head(2)', 'Head-Tail Rate(2)', 'Parallel Rate(2)', 'Head-Head Rate(2)']]
    wb.sheets['data'].range(31,1).value=titles3  


  

    chain=1
    while chain < 26:
        loop_num=[]
        mediate_num=[]

        i=0
        while i<5:
            row=(26-chain)*13-i
            j=5-i
            while j>0:
                column=chain*13-j+1
                loop_num.append(wb.sheets[sheet_name].range(row,column).value)
                j-=1
            i+=1

        i=0
        while i<7:
            row=(26-chain)*13-i
            column=(chain-1)*13+7+i
            if i<6:
                mediate_num.append(wb.sheets[sheet_name].range(row,column).value)
                mediate_num.append(wb.sheets[sheet_name].range(row,column+1).value)
            else:
                mediate_num.append(wb.sheets[sheet_name].range(row,column).value)
            i+=1
        
        loop_dens=np.sum(list(map(lambda x: x < 1, loop_num)))

        #print(loop_num)
        #print(loop_dens)        

        if loop_dens > 0:
            loop_or_not=1
        elif loop_dens == 0:
            loop_or_not=0


        mediate_dens=np.sum(list(map(lambda y: y < 1, mediate_num)))

        #print(mediate_num)
        #print(mediate_dens)  
        if mediate_dens==0 and loop_dens ==0:
            straight_or_not=1
        else:
            straight_or_not=0
        
        if loop_or_not==0 and straight_or_not==0:
            mediate_or_not=1
        else:
            mediate_or_not=0        
        loop_strength=loop_dens/15

        wb.sheets['data'].range(chain+2,colunm_dis).value=[[chain, loop_or_not, loop_dens, loop_strength, straight_or_not, mediate_or_not]]


        if chain==1:
            pairs_y.append(chain)
            pairs_y.append(chain)
            pairs_y.append(chain)
            pairs_y.append(chain)
            pairs_x.append(2)
            pairs_x.append(5)
            pairs_x.append(6)
            pairs_x.append(21)
            pairs_num+=4

            pairs_nd_y.append(chain)
            pairs_nd_y.append(chain)
            pairs_nd_y.append(chain)
            pairs_nd_y.append(chain)
            pairs_nd_x.append(7)
            pairs_nd_x.append(10)
            pairs_nd_x.append(22)
            pairs_nd_x.append(25)
            pairs_nd_num+=4            
        elif chain>1 and chain<5:
            pairs_y.append(chain)
            pairs_y.append(chain)
            pairs_y.append(chain)
            pairs_x.append(chain+1)
            pairs_x.append(chain+5)
            pairs_x.append(chain+20)
            pairs_num+=3

            pairs_nd_y.append(chain)
            pairs_nd_y.append(chain)
            pairs_nd_y.append(chain)
            pairs_nd_y.append(chain)
            pairs_nd_x.append(chain+4)
            pairs_nd_x.append(chain+6)
            pairs_nd_x.append(chain+19)
            pairs_nd_x.append(chain+21)
            pairs_nd_num+=4       
        elif chain==5:
            pairs_y.append(chain)
            pairs_y.append(chain)
            pairs_x.append(10)
            pairs_x.append(25)
            pairs_num+=2

            pairs_nd_y.append(chain)
            pairs_nd_y.append(chain)
            pairs_nd_y.append(chain)
            pairs_nd_y.append(chain) 
            pairs_nd_x.append(6)                    
            pairs_nd_x.append(9)
            pairs_nd_x.append(21)
            pairs_nd_x.append(24)
            pairs_nd_num+=4                     
        elif chain==6 or chain==11 or chain==16:
            pairs_y.append(chain)
            pairs_y.append(chain)
            pairs_y.append(chain)            
            pairs_x.append(chain+1)
            pairs_x.append(chain+4)
            pairs_x.append(chain+5)
            pairs_num+=3

            pairs_nd_y.append(chain)
            pairs_nd_y.append(chain)
            pairs_nd_x.append(chain+6)
            pairs_nd_x.append(chain+9)
            pairs_nd_num+=2               
        elif chain==10 or chain==15 or chain==20:
            pairs_y.append(chain)
            pairs_x.append(chain+5)
            pairs_num+=1

            pairs_nd_y.append(chain)
            pairs_nd_y.append(chain)
            pairs_nd_x.append(chain+1)
            pairs_nd_x.append(chain+4)
            pairs_nd_num+=2             
        elif chain==21:
            pairs_y.append(chain)
            pairs_y.append(chain)
            pairs_x.append(chain+1)
            pairs_x.append(chain+4) 
            pairs_num+=2
        elif chain>21 and chain<25:
            pairs_y.append(chain)
            pairs_x.append(chain+1)
            pairs_num+=1
        elif chain==25:
            pass
        else:
            pairs_y.append(chain)
            pairs_y.append(chain)
            pairs_x.append(chain+1)
            pairs_x.append(chain+5)     
            pairs_num+=2    

            pairs_nd_y.append(chain)
            pairs_nd_y.append(chain)
            pairs_nd_x.append(chain+4)
            pairs_nd_x.append(chain+6)
            pairs_nd_num+=2     



        if chain==1 or chain==2 or chain==6 or chain==7:
            pairs_rd_y.append(chain)
            pairs_rd_y.append(chain)
            pairs_rd_y.append(chain)
            pairs_rd_y.append(chain)
            pairs_rd_x.append(chain+2)
            pairs_rd_x.append(chain+3)
            pairs_rd_x.append(chain+10)
            pairs_rd_x.append(chain+15)
            pairs_rd_num+=4            
        elif chain==3 or chain==8:
            pairs_rd_y.append(chain)
            pairs_rd_y.append(chain)
            pairs_rd_y.append(chain)
            pairs_rd_x.append(chain+2)
            pairs_rd_x.append(chain+10)
            pairs_rd_x.append(chain+15)
            pairs_rd_num+=3
        elif chain==4 or chain==5 or chain==9 or chain==10:
            pairs_rd_y.append(chain)
            pairs_rd_y.append(chain)
            pairs_rd_x.append(chain+10)
            pairs_rd_x.append(chain+15)
            pairs_rd_num+=2                                
        elif chain==11 or chain==12:
            pairs_rd_y.append(chain)
            pairs_rd_y.append(chain)
            pairs_rd_y.append(chain)            
            pairs_rd_x.append(chain+2)
            pairs_rd_x.append(chain+3)
            pairs_rd_x.append(chain+10)            
            pairs_rd_num+=3                
        elif chain==13:
            pairs_rd_y.append(chain)
            pairs_rd_y.append(chain)
            pairs_rd_x.append(15)
            pairs_rd_x.append(23)
            pairs_rd_num+=2             
        elif chain==14 or chain==15:
            pairs_rd_y.append(chain)
            pairs_rd_x.append(chain+10)
            pairs_rd_num+=1
        elif chain==16 or chain==17 or chain==21 or chain==22:
            pairs_rd_y.append(chain)
            pairs_rd_y.append(chain)            
            pairs_rd_x.append(chain+2)
            pairs_rd_x.append(chain+3)            
            pairs_rd_num+=2
        elif chain==18 or chain==23:
            pairs_rd_y.append(chain)
            pairs_rd_x.append(chain+2)
            pairs_rd_num+=1
        else:
            pass

        chain+=1   


    loop_sum=np.sum(wb.sheets['data'].range((3,colunm_dis+1),(27,colunm_dis+1)).value)
    straight_sum=np.sum(wb.sheets['data'].range((3,colunm_dis+4),(27,colunm_dis+4)).value)
    mediate_sum=np.sum(wb.sheets['data'].range((3,colunm_dis+5),(27,colunm_dis+5)).value)

    loop_per=loop_sum/25
    straight_per=straight_sum/25
    mediate_per=mediate_sum/25    

    wb.sheets['data'].range(28,colunm_dis).value=[['Sum', loop_sum, '', '', straight_sum, mediate_sum]]
    wb.sheets['data'].range(29,colunm_dis).value=[['Percent', loop_per, '', '', straight_per, mediate_per]]    

    print(pairs_y)
    print(pairs_x)
    print(pairs_nd_y)
    print(pairs_nd_x)   
    print(pairs_rd_y)
    print(pairs_rd_x)      
    print('pairs_num')
    print(pairs_num)
    print('pairs_nd_num')
    print(pairs_nd_num)
    print('pairs_rd_num')
    print(pairs_rd_num)    

    CHAIN_CHAIN(_dis,1,sheet_name,pairs_y,pairs_x,pairs_num)
    CHAIN_CHAIN(_dis,1.414,sheet_name,pairs_nd_y,pairs_nd_x,pairs_nd_num)
    CHAIN_CHAIN(_dis,2,sheet_name,pairs_rd_y,pairs_rd_x,pairs_rd_num)


    wb.sheets['data'].range('A1:LP326').columns.autofit() 

    wb.sheets['data'].range(row_dis,1).value=[[grafting_dens, _dis, loop_sum, straight_sum, mediate_sum, loop_per, straight_per, mediate_per, '']]

#############################################################################
##########MAIN####MAIN########MAIN#######MAIN############MAIN################
#############################################################################



DISTANCE=[1.0,1.2,1.34,1.62,2.0]

SALT=[0]


PATH='F:\\simulation\\3.Zwitterionic-Peptide-Brush\\Brush\\20200822\\CEEEEEEKKKKKK\\'
#PATH='F:\\simulation\\3.Zwitterionic-Peptide-Brush\\Brush\\20200822\\CEEEKKKEEEKKK\\'
#PATH='F:\\simulation\\3.Zwitterionic-Peptide-Brush\\Brush\\20200822\\CEEKKEEKKEEKK\\'
#PATH='F:\\simulation\\3.Zwitterionic-Peptide-Brush\\Brush\\20200822\\CEKEKEKEKEKEK\\'
#PATH='F:\\simulation\\3.Zwitterionic-Peptide-Brush\\Brush\\20200822\\CKEKEKEKEKEKE\\'
#PATH='F:\\simulation\\3.Zwitterionic-Peptide-Brush\\Brush\\20200822\\CKKEEKKEEKKEE\\'
#PATH='F:\\simulation\\3.Zwitterionic-Peptide-Brush\\Brush\\20200822\\CKKKEEEKKKEEE\\'
#PATH='F:\\simulation\\3.Zwitterionic-Peptide-Brush\\Brush\\20200822\\CKKKKKKEEEEEE\\'

SEQUENCE = 'E6K6'
#SEQUENCE = 'E3K3'
#SEQUENCE = 'E2K2'
#SEQUENCE = 'EK'
#SEQUENCE = 'KE'
#SEQUENCE = 'K2E2'
#SEQUENCE = 'K3E3'
#SEQUENCE = 'K6E6'



app=xw.App(visible=True,add_book=False)    
wb = app.books.open(PATH+'map-'+SEQUENCE+'.xlsx')


wb.sheets.add('data')

for i in range(len(SALT)):
    for j in range(len(DISTANCE)):       
        if SALT[i] == 0:
            file_path = PATH+str(DISTANCE[j])+'\\'+str(SALT[i])+'\\'
        else:
            file_path = PATH+str(DISTANCE[j])+'\\'+str(SALT[i])+'_NACL\\'       
        DATALOAD(DISTANCE[j],SALT[i])
        ANALYZE(DISTANCE[j],SALT[i])
#############################################################################
#############################################################################



#wb.save('data.xlsx')

wb.save()
wb.close()
app.quit()


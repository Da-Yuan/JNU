clear;clc;


%起始日期
start_date = datenum('2017-02-28');
%截止日期
end_date =   datenum('2017-03-01');
%要读取的文件名
read_filename = '宣讲会信息导出.xlsx';



%读取文件
[NUM,TXT,RAW]=xlsread(read_filename);
datanum = size(TXT,2);
time_list='';
date_diff =zeros(datanum,1);
T(1,:) = {'日期','预约时间','公司名称','宣讲地址','联系人','电话','手机','笔试时间','笔试地址','面试时间','面试地址','下发学院'};
j=2;
for i=2:size(TXT,1)
    a=TXT(i,5);
    a=char(a);
    time_list(i-1,:)=a(1:10);
    date_diff(i-1,1) = datenum(a(1:10),'yyyy-mm-dd') - start_date;%>=0
    date_diff(i-1,2) = datenum(a(1:10),'yyyy-mm-dd') - end_date;%<=0
    date_diff(i-1,3) = i-1;
    
    if(date_diff(i-1,1)>=0  &&  date_diff(i-1,2)<=0)
        % 日期	
        date = TXT{i,5};
        T{j,1} =  date(1:10);
        % 预约时间
        T{j,2} =  date(11:end);
        % 公司名称
        T(j,3) =  TXT(i,1);
        % 宣讲地址
        T(j,4) =  TXT(i,6);
        % 联系人
        T(j,5) =  TXT(i,15);
        % 电话
        T(j,6) =  TXT(i,16);
        % 手机
        T(j,7) =  TXT(i,17);
        % 笔试时间
        T(j,8) =  TXT(i,7);
        % 笔试地址
        T(j,9) =  TXT(i,8);
        % 面试时间
        T(j,10) =  TXT(i,9);
        % 面试地址
        T(j,11) =  TXT(i,10);
        % 下发学院
        T(j,12) =  TXT(i,19);       
   
        j = j + 1;
    end
    
end
filename = [datestr(start_date,'mmdd') '-' datestr(end_date,'mmdd') '.xlsx']
xlswrite(filename,T);
display(['处理完成,已保存为' filename])

    

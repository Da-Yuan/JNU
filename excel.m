clear;clc;


%��ʼ����
start_date = datenum('2017-02-28');
%��ֹ����
end_date =   datenum('2017-03-01');
%Ҫ��ȡ���ļ���
read_filename = '��������Ϣ����.xlsx';



%��ȡ�ļ�
[NUM,TXT,RAW]=xlsread(read_filename);
datanum = size(TXT,2);
time_list='';
date_diff =zeros(datanum,1);
T(1,:) = {'����','ԤԼʱ��','��˾����','������ַ','��ϵ��','�绰','�ֻ�','����ʱ��','���Ե�ַ','����ʱ��','���Ե�ַ','�·�ѧԺ'};
j=2;
for i=2:size(TXT,1)
    a=TXT(i,5);
    a=char(a);
    time_list(i-1,:)=a(1:10);
    date_diff(i-1,1) = datenum(a(1:10),'yyyy-mm-dd') - start_date;%>=0
    date_diff(i-1,2) = datenum(a(1:10),'yyyy-mm-dd') - end_date;%<=0
    date_diff(i-1,3) = i-1;
    
    if(date_diff(i-1,1)>=0  &&  date_diff(i-1,2)<=0)
        % ����	
        date = TXT{i,5};
        T{j,1} =  date(1:10);
        % ԤԼʱ��
        T{j,2} =  date(11:end);
        % ��˾����
        T(j,3) =  TXT(i,1);
        % ������ַ
        T(j,4) =  TXT(i,6);
        % ��ϵ��
        T(j,5) =  TXT(i,15);
        % �绰
        T(j,6) =  TXT(i,16);
        % �ֻ�
        T(j,7) =  TXT(i,17);
        % ����ʱ��
        T(j,8) =  TXT(i,7);
        % ���Ե�ַ
        T(j,9) =  TXT(i,8);
        % ����ʱ��
        T(j,10) =  TXT(i,9);
        % ���Ե�ַ
        T(j,11) =  TXT(i,10);
        % �·�ѧԺ
        T(j,12) =  TXT(i,19);       
   
        j = j + 1;
    end
    
end
filename = [datestr(start_date,'mmdd') '-' datestr(end_date,'mmdd') '.xlsx']
xlswrite(filename,T);
display(['�������,�ѱ���Ϊ' filename])

    

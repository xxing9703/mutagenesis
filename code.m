%*************************************************************************
% this code is used to modify genes based on condon tables, and export to
% excels files
% required files: 1) condon1, 2) condon2, 3)parent, 4)project
%*************************************************************************

clear
%con1=readtable('condon1a.txt','ReadVariableNames',false); %lower case, load condon table 1-1 for modifications 
con1=readtable('condon1.txt','ReadVariableNames',false); %upper case, load condon table 1-1 for modifications
con2=readtable('condon2.txt','ReadVariableNames',false);  %load condon table, for validation
c1=table2cell(con1);
c2=table2cell(con2);

p=textread('parent2.txt','%c');  %original sequence
p0=textread('parent00.txt','%c'); 
%-----------------------------------
for i=1:length(p)/3
 cc{i}=p(i*3-2:i*3)';
 id1=strmatch(cc{i}, char(c2(:,2)));
 letter{i}=c2{id1,1};
 idx{i}=int2str(i);
 
 W{1,i+1}=idx{i};  %W row 1
W{2,i+1}=letter{i}; %W row 2
W{3,i+1}=cc{i}; %W row 3
end




%----------------------------------------------------modify tasks below
f=fopen('tab2.txt');   %open task list, multiple ones using '/' to separate in the same line
delimiter='/';

tline = fgetl(f);  %start read line by line
disp(tline)

count=0; %total entries
clrcount=0;
while ischar(tline)
    count=count+1;
    tr=textscan(tline,'%s','Delimiter',delimiter);
    tr=tr{1}; 
    s=size(tr,1);
    output=p;
for i=1:s
   AA=textscan(tr{i},'%c %d %c'); %parse, into letter, number, letter, store in AA{1} AA{2} AA{3}
   
   id=strmatch(AA{3}, char(c1(:,1)));   %look up AA{3} in condon table 1
   char2T=c1{id,2};  %find the corresponding gene in table
   
   check=output(AA{2}*3-2:AA{2}*3);  %check if AA{1} is correct
   id1=strmatch(check, char(c2(:,2)));
   char1=AA{1};
   char2=c2{id1,1};
   if ~strcmp(char1(1),char2(1))  % report error
       fprintf('*****error')
       [char1 char2]
       AA{2}
   end
   
   output(AA{2}*3-2)=char2T(1);  %make modifications
   output(AA{2}*3-1)=char2T(2);
   output(AA{2}*3-0)=char2T(3);   
   clrcount=clrcount+1;
   clr(clrcount,1)=count+3;
   clr(clrcount,2)=AA{2}+1;
end
    O{count,1}=tline;    %save to table O
  O{count,2}=output';
  
  W{count+3,1}=tline;   %save to extended table W
  
  for i=1:length(letter)
    W{count+3,i+1}=output(i*3-2:i*3)';
  end
  
  
  tline = fgetl(f);  %  read next line
  disp(tline)
end

savefile='tab2cl.xlsx';
xlswrite(savefile,W)    % export to excel file

Excel = actxserver('excel.application');
WB = Excel.Workbooks.Open(fullfile(pwd, savefile),0,false);

for i=1:clrcount   %find the cells, mark the color
 rg=strcat(ExcelCol(clr(i,2)),int2str(clr(i,1)));  
rg=rg{1};
 WB.Worksheets.Item(1).Range(rg).Interior.ColorIndex = 4;
 WB.Save();
end

WB.Close();
Excel.Quit();


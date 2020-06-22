function ExtractData4Excel
% WriteWordTable write the table in Word Document
% ruogu7(380545156@qq.com)
% Start: Am 8:14 Feb. 14th, 2020
% End: Pm 8:14 Feb. 14th, 2020
%
% Parameters:
% file_Path: the path of template file;
% copy_Num: the num of copies;
%
% Example:
% ExtractData4Excel( 'data_people.xlsx')
%
clc; clear all;
% extract the data from excel file
[numdata,textdata,alldata] = xlsread('data_people.xlsx');
String_alldata = string(alldata);
Cleandata = String_alldata(2:end,:);

kkk = unique(str2double(Cleandata(:,1)'))

for ii = 1:size(kkk,2)
    Pre_name = int2str(kkk(ii));
    % Copy a file with the template
    Object_filename = [Pre_name,'.docx'];
    CopyWORD_Template( Object_filename)
    
    % Extract the data and write in the word table
    data4table = Cleandata((str2double(Cleandata(:,1)')')== kkk(ii),:);
    
    %% open the doc
    word = actxserver('Word.Application');
    % filename_in =  '28.docx';
    pathname =  'D:\郑洲\郑州市_不动产\save2word\';
    document = word.documents.Open(fullfile(pathname,Object_filename));
    % tbl_cnt = document.Tables.Count
    % row_cnt = document.Tables.Item(1).Rows.Count
    % col_cnt = document.Tables.Item(1).Columns.Count
    % cell_txt1 = strtrim(document.Tables.Item(1).Cell(3,2).Range.Text)
    % cell_txt2 = strtrim(document.Tables.Item(1).Cell(3,4).Range.Text)
    % cell_txt3 = strtrim(document.Tables.Item(1).Cell(4,2).Range.Text)
    % cell_txt4 = strtrim(document.Tables.Item(1).Cell(4,4).Range.Text)
    
    document.Tables.Item(1).Cell(3,2).Range.Text = data4table(1,1)   % 宗地号
    document.Tables.Item(1).Cell(3,4).Range.Text = data4table(1,2)   % 户主姓名
    document.Tables.Item(1).Cell(4,2).Range.Text = data4table(1,4)   % 联系电话
    document.Tables.Item(1).Cell(4,4).Range.Text = data4table(1,7)   % 土地坐落
    %     document.Tables.Item(1).Cell(3,2).Range.Text = 'Zongdihao' % 宗地号
    %     document.Tables.Item(1).Cell(3,4).Range.Text = 'HuZhu'  % 户主姓名
    %     document.Tables.Item(1).Cell(4,2).Range.Text = 'LianXiDianHua' % 联系电话
    %     document.Tables.Item(1).Cell(4,4).Range.Text = 'TuDiZuoLuo'  % 土地坐落
    
    for jjj = 1:size(data4table,1)
        document.Tables.Item(1).Cell(12+jjj,1).Range.Text = data4table(jjj,2)   % 姓名
        document.Tables.Item(1).Cell(12+jjj,2).Range.Text = data4table(jjj,3)     % 身份证
        document.Tables.Item(1).Cell(12+jjj,3).Range.Text = data4table(jjj,5)     % 与户主关系
    end
    % Save the document.
    document.Save;
    % Close the document.
    document.Close;
    % Close Word.
    word.Quit;
    % delete server object
    delete(word);
    clear word document    
end
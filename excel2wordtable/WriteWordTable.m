function WriteWordTable( WordwithTable  )
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
% WriteWordTable( 'template_original.docx',8)
%
% Steps:
% 
clc; clear all;
%% open the doc
word = actxserver('Word.Application');
% Open/Select Word file
% [filename_in, pathname] = uigetfile({'*.doc;*.docx;*.docm','Word Files (*.doc,*.docx,*.docm)'; ...
%     '*.*', 'All Files (*.*)'}, 'Select a file')

filename_in =  '28.docx';
pathname =  'C:\D_Development\郑洲\郑州市_不动产\save2word\';
document = word.documents.Open(fullfile(pathname,filename_in));
tbl_cnt = document.Tables.Count
row_cnt = document.Tables.Item(1).Rows.Count
col_cnt = document.Tables.Item(1).Columns.Count

cell_txt1 = strtrim(document.Tables.Item(1).Cell(3,2).Range.Text)
cell_txt2 = strtrim(document.Tables.Item(1).Cell(3,4).Range.Text)
cell_txt3 = strtrim(document.Tables.Item(1).Cell(4,2).Range.Text)
cell_txt4 = strtrim(document.Tables.Item(1).Cell(4,4).Range.Text)

document.Tables.Item(1).Cell(3,2).Range.Text = 'Zongdihao' % 宗地号
document.Tables.Item(1).Cell(3,4).Range.Text = 'HuZhu'  % 户主姓名
document.Tables.Item(1).Cell(4,2).Range.Text = 'LianXiDianHua' % 联系电话
document.Tables.Item(1).Cell(4,4).Range.Text = 'TuDiZuoLuo'  % 土地坐落

% Save the document.
document.Save;
% Close the document.
document.Close;
% Close Word.
word.Quit;
% delete server object
delete(word);
clear word document

return
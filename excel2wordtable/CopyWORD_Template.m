function CopyWORD_Template(filename)

% CopyWORD_Template copys the WORD Template with copy_Num copies
% ruogu7(380545156@qq.com) 
% Start: Am 8:12 Feb. 11th, 2020
% End: Am 8:12 Feb. 11th, 2020
%
% Parameters:
% file_Path: the path of template file;
% copy_Num: the num of copies;
%
% Example: 
% CopyWORD_Template( '500.docx')
%
% Steps:
% 
% %% open the doc
% word = actxserver('Word.Application');
% % Open/Select Word file
% [filename_in, pathname] = uigetfile({'*.doc;*.docx;*.docm','Word Files (*.doc,*.docx,*.docm)'; ...
%     '*.*', 'All Files (*.*)'}, 'Select a file')
% 
% document = word.documents.Open(fullfile(pathname,filename_in));

file_original = 'template_original.docx';
% filename = [Pre_filename,'.docx'];
status = copyfile(file_original,filename);

return
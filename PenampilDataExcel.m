function varargout = PenampilDataExcel(varargin)
% PENAMPILDATAEXCEL MATLAB code for PenampilDataExcel.fig
%      PENAMPILDATAEXCEL, by itself, creates a new PENAMPILDATAEXCEL or raises the existing
%      singleton*.
%
%      H = PENAMPILDATAEXCEL returns the handle to a new PENAMPILDATAEXCEL or the handle to
%      the existing singleton*.
%
%      PENAMPILDATAEXCEL('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in PENAMPILDATAEXCEL.M with the given input arguments.
%
%      PENAMPILDATAEXCEL('Property','Value',...) creates a new PENAMPILDATAEXCEL or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before PenampilDataExcel_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to PenampilDataExcel_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help PenampilDataExcel

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @PenampilDataExcel_OpeningFcn, ...
                   'gui_OutputFcn',  @PenampilDataExcel_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT
% --- Executes during object creation, after setting all properties.


% --- Executes just before PenampilDataExcel is made visible.
function PenampilDataExcel_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to PenampilDataExcel (see VARARGIN)

% Choose default command line output for PenampilDataExcel
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);
global s;
s = serial('COM3','baudrate',9600);

%edit text%
function edit2_Callback(hObject, eventdata, handles)


% --- Outputs from this function are returned to the command line.
function varargout = PenampilDataExcel_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;
imshow('UIN.png', 'Parent', handles.axes3);



% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global s;
fopen(s);
i = 0;
pr = guidata(gcbo);
userInput = str2double (get(pr.edit2, 'string'));
while (i < userInput)
    i = i + 1;
    c=(clock);
    w(i) = c(5); %menit
    z(i) = c(6); %detik
    u(i) = i;
    y(i) = fscanf(s,'%d');
    fh = ((9/5)*y(i))+32;
    rr = (4/5)* y(i);
    kv = y(i);
    set(handles.text3,'string',num2str(i));
    set(handles.text12,'string',num2str(kv));
    set(handles.text4,'string',num2str(y(i)));
    set(handles.text14,'string',num2str(fh));
    set(handles.text16,'string',num2str(rr));
    axes(handles.axes1);
    plot(y,'r.-','LineWidth',2);
    grid on;
    axis([0 userInput 0 120]);
end
fclose(s);
xlabel('Waktu (Sekon)');
ylabel('Suhu (C)');
print -dmeta;
filename = strcat('D:\',num2str(c(3)),num2str(c(4)),num2str(c(5)),'.xlsx');
xlswrite(filename,{'Hasil Pengukuran'},'Sheet1','A1');
xlswrite(filename,{'waktu','Suhu (C)'},'Sheet1','A2');
xlswrite(filename,[u;y]','Sheet1','A3');
Excel = actxserver('Excel.Application');
Excel.Visible = 1;
invoke(Excel.Workbooks,'Open',filename);
ActiveSheet = Excel.ActiveSheet;
ActiveSheetRange = get(ActiveSheet,'Range','F2');
ActiveSheetRange.Select;
ActiveSheetRange.PasteSpecial;



% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
[file,path] = uigetfile('*.xlsx');
fileku = fullfile(path,file);
handles.fileku = fileku;
[numbers,colNames] = xlsread(fileku,'Sheet1','2:2');
set(handles.popupmenu1,'string',colNames);
set(handles.popupmenu2,'string',colNames);
set(handles.popupmenu1,'callback','PenampilDataExcel(''updateSumbu'',gcbo,[],guidata(gcbo))');
set(handles.popupmenu2,'callback','PenampilDataExcel(''updateSumbu'',gcbo,[],guidata(gcbo))');
guidata(hObject,handles);

function updateSumbu(hObject, eventdata, handles)
fileku = handles.fileku;
xColNum = get(handles.popupmenu1,'value');
yColNum = get(handles.popupmenu2,'value');
dataku = xlsread(fileku);
x = dataku(:,xColNum);
y = dataku(:,yColNum);
plot(handles.axes1,x,y);




% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu1 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu1


% --- Executes during object creation, after setting all properties.
function popupmenu1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu2.
function popupmenu2_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu2 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu2


% --- Executes during object creation, after setting all properties.
function popupmenu2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



% --- Executes during object creation, after setting all properties.
function edit2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes during object creation, after setting all properties.
function text4_CreateFcn(hObject, eventdata, handles)
% hObject    handle to text4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called


% --- Executes during object deletion, before destroying properties.
function text4_DeleteFcn(hObject, eventdata, handles)
% hObject    handle to text4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)




% Hint: place code in OpeningFcn to populate axes2


% --- Executes during object creation, after setting all properties.
function axes3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to axes3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: place code in OpeningFcn to populate axes3

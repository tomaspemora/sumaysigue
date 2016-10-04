function varargout = sumaysigue(varargin)
% SUMAYSIGUE MATLAB code for sumaysigue.fig
%      SUMAYSIGUE, by itself, creates a new SUMAYSIGUE or raises the existing
%      singleton*.
%
%      H = SUMAYSIGUE returns the handle to a new SUMAYSIGUE or the handle to
%      the existing singleton*.
%
%      SUMAYSIGUE('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in SUMAYSIGUE.M with the given input arguments.
%
%      SUMAYSIGUE('Property','Value',...) creates a new SUMAYSIGUE or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before sumaysigue_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to sumaysigue_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help sumaysigue

% Last Modified by GUIDE v2.5 08-Sep-2016 20:42:24

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @sumaysigue_OpeningFcn, ...
                   'gui_OutputFcn',  @sumaysigue_OutputFcn, ...
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

% --- Executes just before sumaysigue is made visible.
function sumaysigue_OpeningFcn(hObject, eventdata, handles, varargin)
warning off;
% Choose default command line output for sumaysigue
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% Icono Suma y Sigue en la ventana
javaFrame = get(hObject,'JavaFrame');
javaFrame.setFigureIcon(javax.swing.ImageIcon('icon.jpg'));
warning on;

% This sets up the initial plot - only do when we are invisible
% so window can get raised using sumaysigue.


% UIWAIT makes sumaysigue wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = sumaysigue_OutputFcn(hObject, eventdata, handles)
varargout{1} = handles.output;

function CloseMenuItem_Callback(hObject, eventdata, handles)
selection = questdlg(['Close ' get(handles.figure1,'Name') '?'],...
                     ['Close ' get(handles.figure1,'Name') '...'],...
                     'Yes','No','Yes');
if strcmp(selection,'No')
    return;
end
delete(handles.figure1)

function Ayuda_Callback(hObject, eventdata, handles)

function Acerca_Callback(hObject, eventdata, handles)

function pushbutton2_Callback(hObject, eventdata, handles)
[filename ,pathname, filterIndex] = uigetfile({'*.xls;*.xlsx;','Archivos Excel'}, 'Seleccione el archivo de plantilla');
if filename == 0
    hObject.String = 'seleccionar archivo...';    
else
    if length(pathname) <= 25
        hObject.String = ['...' pathname filename];
    else
        hObject.String = ['...' pathname(end-25:end) filename];
    end
    hObject.UserData.files = filename;
    hObject.UserData.path = pathname;
end

function Configuracion_Callback(hObject, eventdata, handles)
configuration;

function pushbutton4_Callback(hObject, eventdata, handles)
[filename ,pathname, filterIndex] = uigetfile({'*.xls;*.xlsx;','Archivos Excel'}, 'Seleccione el o los archivos con las respuestas', 'MultiSelect', 'on');
if ~iscell(filename) & filename == 0
    hObject.String = 'seleccionar archivo(s)...';    
else
    if iscell(filename)
        if length(pathname) <= 20
            hObject.String = ['...' pathname '*varios archivos*'];
        else
            hObject.String = ['...' pathname(end-20:end) '*varios archivos*'];
        end
    else
        if length(pathname) <= 25
            hObject.String = ['...' pathname filename];
        else
            hObject.String = ['...' pathname(end-25:end) filename];
        end
    end
    hObject.UserData.files = filename;
    hObject.UserData.path = pathname;
end

function pushbutton5_Callback(hObject, eventdata, handles)

try
    
    string_length_console = 90;                                            % LARGO DE STRING EN CONSOLA
    COMP = computer;
    
    %% EXTRACCION DE ARCHIVOS DE ENTRADA DE USUARIO
    r_file          = handles.pushbutton2.UserData.files;                   % ARCHIVO DE PLANTILLA DE REIVISION
    r_path          = handles.pushbutton2.UserData.path;                    % RUTA DE ARCHIVO DE PLANTILLA DE REVISION
    a_file          = cellstr(handles.pushbutton4.UserData.files);          % ARCHIVO(S) DE RESPUESTAS
    a_path          = handles.pushbutton4.UserData.path;                    % RUTA DE ARCHIVO(S) DE RESPUESTAS
    
    %% EXTRACCION DE PARAMETROS DE CONFIGURACION
    parametros      = xml2struct('datos.xml');                              % ARCHIVO DE DATOS DE CONFIGURACION
    fl_data         = str2double(parametros.data.fl_data.Text);             % FILA EN QUE EMPIEZA LA DATA DE PLANTILLA
    r_col_rut       = str2double(parametros.data.r_col_rut.Text);           % INDICE COLUMNA DE RUT EN PLANTILLA
    r_col_data_1    = str2double(parametros.data.r_col_data_1.Text);        % INDICE COLUMNA EN QUE TERMINAN LOS DATOS DE PERSONAS EN PLANTILLA
    r_col_data_2    = str2double(parametros.data.r_col_data_2.Text);        % INDICE COLUMNA EN QUE EMPIEZAN LAS RESPUESTAS EN PLANTILLA
    N               = str2double(parametros.data.N.Text);                   % NUMERO DE ALUMNOS
    rev_sheet       = parametros.data.pes_rev.Text;                         % PESTAÑA DE PLANTILLA DE REVISION
    abi_sheet       = parametros.data.pes_abi.Text;                         % PESTAÑA DE PLANTILLA DE PREGUNTAS ABIERTAS
    string_flag     = parametros.data.string_flag.Text;                     % STRING PARA IDENTIFICAR LAS PREGUNTAS
    %a_col_nombre    = 2;                                                    % INDICE COLUMNA DE NOMBRE EN ARCHIVOS RESPUESTAS 
    a_col_fecha     = 3;                                                    % INDICE COLUMNA DE FECHA EN ARCHIVOS RESPUESTAS
    a_col_rut       = str2double(parametros.data.a_col_rut.Text);           % INDICE COLUMNA DE RUT EN ARCHIVOS RESPUESTAS
    fl_a_data       = str2double(parametros.data.fl_a_data.Text);           % FILA EN QUE EMPIEZAN LAS RESPUESTAS EN ARCHIVOS DE RESPUESTAS
    info_fil        = 1;                                                    % FILA CON INFO DE ARCHIVOS RESPUESTAS
    ll_data         = N+fl_data-1;                                          % FILA EN QUE TERMINAR LA DATA DE PLANTILLA
    
    switch COMP
        case 'MACI64'
        % CASE MAC
            SheetsStruct    = importdata([r_path r_file]);
            Sheets          = fieldnames(SheetsStruct.textdata);
        case 'PCWIN64'
        % CASE WINDOWS
            [~,Sheets,~]    = xlsfinfo([r_path r_file]);                            % SHEETS: CELL ARRAY CON NOMBRE PESTAÑA EN PLANTILLA (CAMBIAR POR ENTRADA USUARIO)
            Sheets          = cellstr(Sheets);            
        otherwise
        % CASE OTHERS
            log_s = 'Sistema Operativo no soportado.';
            log_s = fillString(log_s,string_length_console);
            handles.text4.String = [handles.text4.String; log_s];
            return;            
    end
        
    %% COMPROBAR QUE PESTAÑAS DE REVISION Y PREGUNTAS ABIERTAS EXISTAN EN ARCHIVO DE REVISION
    bp_rev = 0; bp_abi = 0;
    for sh = 1:length(Sheets)
        %disp([rev_sheet ' -@- ' sheets{sh} ' -@- ' Sheets{sh}])
        bp_rev = strcmp(Sheets{sh},rev_sheet) + bp_rev;
        bp_abi = strcmp(Sheets{sh},abi_sheet) + bp_abi;
    end
    if ~bp_rev
        log_s = ['Revise el nombre de la pestaña de revisión. La pestaña ''' rev_sheet ''' no se encuentra en ''' r_file ''''];
        log_s = fillString(log_s,string_length_console);
        handles.text4.String = [handles.text4.String; log_s];
        return;
    elseif ~bp_abi
        log_s = ['Revise el nombre de la pestaña de preguntas abiertas. La pestaña ''' abi_sheet ''' no se encuentra en ''' r_file ''''];
        log_s = fillString(log_s,string_length_console);
        handles.text4.String = [handles.text4.String; log_s];
        return;
    end
    
    %% ABRIR ARCHIVOS DE REVISION
    switch COMP
        case 'PCWIN64'
        [~,~,Rcop]      = xlsread([r_path r_file],rev_sheet);                   % RCOP: CELL ARRAY CON DATOS PLANTILLA REVISION
        otherwise
        % CASE OTHERS
        log_s = 'Sistema Operativo no soportado.';
        log_s = fillString(log_s,string_length_console);
        handles.text4.String = [handles.text4.String; log_s];
        return;     
    end
    %[~,~,Acop]      = xlsread([r_path r_file],abi_sheet);                   % ACOP: CELL ARRAY CON DATOS PLANTILLA PREGUNTAS ABIERTAS
    stud_data_cop   = Rcop(fl_data:ll_data,1:r_col_data_1-1);               % INFORMACION DE LOS ALUMNOS, NO DEBE TENER MAS DE 4 COLUMNAS, i.e r_col_data_1 = 5 siempre
    stud_data_date  = stud_data_cop;
    course_data     = Rcop(1:fl_data-1,r_col_data_2:end);                   % INDICE DE LOS TALLERES y ACTIVIDADES DEL CURSO
    [~, cdL]        = size(course_data);                                    % NUMERO DE TALLERES Y ACTIVIDADES    

    %% ARREGLO DE INDICE DE TALLERES Y ACTIVIDADES     
    for n = 1:cdL
        if isnan(course_data{1,n}), course_data{1,n} = course_data{1,n-1}; end  % RELLENO DE CAMPOS VACIOS EN INDICE DE TALLERES
        if isnan(course_data{2,n}), course_data{2,n} = course_data{2,n-1}; end  % RELLENO DE CAMPOS VACIOS EN INDICE DE ACTIVIDADES
    end
    
    %% CONVERSION DE INDICE DE TALLERES Y ACTIVIDADES A INDICE NUMERICO
	aux_m_0  = course_data{1,1};                                            % VAR AUX PARA TALLER            
    aux_m_1  = course_data{2,1};                                            % VAR AUX PARA ACTIVIDAD
    cont_0   = 1;                                                           % CONTADOR TALLER
    cont_1   = 1;                                                           % CONTADOR ACTIVIDAD   
    for n = 1:cdL
        aux_0 = course_data{1,n};                                           % TALLER
        aux_1 = course_data{2,n};                                           % ACTIVIDAD
        aux_2 = course_data{3,n};                                           % PAGINA
        aux_3 = course_data{4,n};                                           % PREGUNTA
        aux_4 = 0;                                                          % ABIERTA O NO ABIERTA
        
        if ~strcmp(aux_0,aux_m_0)                                           % COMPRUEBA CAMBIO EN TALLER Y ACTIVIDAD
            cont_0 = cont_0 + 1; cont_1 = 1;
        else   
            if ~strcmp(aux_1,aux_m_1), cont_1 = cont_1 + 1; end
        end
        
        if (isnan(aux_2) & isnan(aux_3)) | (ischar(aux_2) & ischar(aux_3)) | length(aux_2)>9 | length(aux_3)>9  %#ok<OR2,AND2> % COMPRUEBA FIN DE INDICES
            break;
        end     

        if isnan(aux_2),aux_2 = CD(n-1,3);end                               % REPARANDO LAS ACTIVIDADES CON MAS DE UNA PREGUNTA (NaN POR ACT. ANTERIOR)

        if strcmp(aux_3,'-')                                                % VERIFICANDO TIPO DE PREGUNTA ABIERTA O NO ABIERTA
            aux_3 = 0;
        else
            if ischar(aux_3)
                aux_33 = aux_3;
                saux = strsplit(aux_3,'(A)');
                if length(saux) > 1, aux_4 = 1; end
                aux_3 = str2double(saux{1});
                aux_33 = regexprep(aux_33,'[^\0-9]','');        
                if isnan(aux_3), aux_3 = str2double(aux_33); end
            end
        end
        CD(n,:) = [cont_0 cont_1 aux_2 aux_3 aux_4];                        % TALLER | ACTIVIDAD | PAGINA | PREGUNTA | ABIERTA/NOABIERTA
        aux_m_0 = aux_0;
        aux_m_1 = aux_1;
    end

    % RESTRICCIONES
    % TALLER: SOLO SE ADMITE TEXTO PARA LAS CASILLAS, ALFANUMERICO
    % ACTIVIDAD: SOLO SE ADMITE TEXTO PARA LAS CASILLAS ALFANUMERICO
    % PAGINA: SOLO SE ADMITE UN NUMERO PARA LAS CASILLAS
    % PREGUNTA: ADMITE NUMERO O NUMERO Y LETRAS O "-": ej: 1, 2, 3a , 3b, 4(A), -

    %% REVISANDO TODOS LOS ARCHIVOS DE RESPUESTAS SELECCIONADOS
    for z = 1:length(a_file)
                
        log_s = ['Revisando el archivo ' a_file{z} ' ...'];
        log_s = fillString(log_s,string_length_console);
        handles.text4.String = [handles.text4.String; log_s];
        
        clear P OA
        outArr                  = stud_data_cop;                                % VARIABLE DE COPIA DE INFORMACION DE LOS ALUMNOS         
        codigo                  = a_file{z};                                    % NOMBRE CODIGO DE ARCHIVO DE RESPUESTAS ej: "variables TXAY_ZZ.xls"
        idx                     = regexp(codigo,'t[\0-9]a[\0-9][_\0-9]');      	% INDICE DE DONDE APARECE TXAY_ZZ.xl...
        codigo                  = strrep(regexprep(codigo(idx:end),'\W\w{1,10}',''),'_','');  % CODIGO LIMPIO TXAYZZ
        switch COMP
            case 'PCWIN64'
                [~,~,R1]                = xlsread([a_path a_file{z}]);                  % CELL ARRAY CON INFO DE ARCHIVO DE RESPUESTAS
            otherwise
                 % CASE OTHERS
                log_s = 'Sistema Operativo no soportado.';
                log_s = fillString(log_s,string_length_console);
                handles.text4.String = [handles.text4.String; log_s];
                return;  
        end
        
        info                    = R1(info_fil,:);                               % PRIMERA FILA DE ARCHIVO DE RESPUESTAS. CONTIENE SUS PARAMETROS
        data                    = R1((info_fil+1):end,:);                       % DATA DE ARCHIVO DE RESPUESTSAS
        I                       = length(info);                                 % NUMERO COLUMNAS ARCHIVO DE RESPUESTAS
        [D,~]                   = size(data);                                   % NUMERO DE ENTRADAS DE ARCHIVO DE RESPUESTAS (FILAS)
        outArr(:,r_col_rut)     = strrep(strrep(strrep(strrep(strrep(...               % REPARACION DE CASILLA DE RUT (-K,-k -> -0 y borrar .)
            outArr(:,r_col_rut),'-',''),'k','0'),'K','0'),'.',''),' ','');
        outArrAB                = outArr;                                       % COPIA PARA ARREGLO DE RESPUESTAS ABIERTAS
        stud_data_date(:,r_col_rut)     = strrep(strrep(strrep(strrep(strrep(...       % REPARACION DE CASILLA DE RUT (-K,-k -> -0 y borrar .)
            stud_data_date(:,r_col_rut),'-',''),'k','0'),'K','0'),'.',''),' ','');

        for i = fl_a_data:I                                                     % CONSTRUCCION DE INDICE DE PREGUNTAS PRESENTES EN ARCHIVOS DE RESPUESTAS
            P(i-fl_a_data+1,:)  = str2num(strrep(strrep(info{i},string_flag,''),'_',' '));
        end
        [pL,~]                  = size(P);                                      % NUMERO DE PREGUNTAS EN ARCHIVO RESPUESTAS

        %% EXTRACCION Y SINTESIS DE RESPUESTAS DESDE ARCHIVO DE RESPUESTAS (FOR SOBRE ENTRADAS (FILAS) DE ARCHIVO DE RESPUESTAS)
        for i = 1:D
            rut     = data{i,a_col_rut};
            rut     = strrep(rut,'-','');
            rut     = strrep(rut,'K','0');
            rut     = strrep(rut,'k','0');
            rut     = strrep(rut,'.','');
            rut     = str2num(rut);     %#ok<*ST2NM>
            %nombre  = data{i,a_col_nombre};
            fecha   = data{i,a_col_fecha};

            for j = fl_a_data:(pL+fl_a_data-1)  % FOR SOBRE PREGUNTAS DE ARCHIVO DE RESPUESTAS (COLUMNAS)
                resp = data{i,j};
                if ~isempty(resp)
                    resp = strsplit(resp,'#');
                    ab_resp = resp{3};
                    resp = strsplit(resp{1}, ':');
                    s    = resp{1};
                    switch s
                        case 'buena'
                            val = 1;
                        case 'mala'
                            val = 0;
                        case 'intento'
                            val = NaN;
                        case 'enviada'
                            val = 2;                            
                    end
                    if ~isnan(val)
                        outArr = addResp(outArr,rut,val,j-fl_a_data+1,r_col_rut,r_col_data_1,handles,string_length_console);
                        if val == 2
                            outArrAB = addResp(outArrAB,rut,ab_resp,j-fl_a_data+1,r_col_rut,r_col_data_1,handles,string_length_console);
                        end
                    end
                    break;
                end
            end
            stud_data_date = addResp(stud_data_date,rut,fecha,fl_a_data+1,r_col_rut,0,handles,string_length_console);
        end
        
        OA(:,1:(fl_a_data-1)) = outArr(:,1:(fl_a_data-1));

        %% POST PROCESS
        NP = max(P(:,1));
        for i = 1:NP
            clear OAC
            C = P(:,1);
            G = find(~(C-i));
            try
                aux = outArr(:,4+(G(1):G(end)));    
                [faux,caux] = size(aux);
                Iaux = isequalCellArray(aux,[]);
                for j = 1:faux
                    for g = 1:caux
                        if Iaux(j,g)
                            aux{j,g} = [NaN];
                        end
                    end
                end
                OAC = cell2mat(aux);
                if length(G)>1
                    OAC = prod(OAC,2);
                end

                OA(:,i+4) = mat2cell(OAC,ones(1,length(OAC)),1);
            catch err
                log_s = ['    La pregunta ' num2str(i) ' no se respondía! OJO!'];
                log_s = fillString(log_s,string_length_console);
                handles.text4.String = [handles.text4.String; log_s];            
                log_me = fillString(['    ' err.message],string_length_console);
                handles.text4.String = [handles.text4.String; log_me];

            end  
        end
        codigo_page = codigo(end-1:end);
        codigo(end-1:end) = [];
        c_codigo = str2double(strsplit(codigo,'[/t/a]','DelimiterType','RegularExpression'));
        c_codigo(1) = [];
        c_codigo(3) = str2double(codigo_page);

        [cdL,~] = size(CD);
        CR = ones(cdL,1)*c_codigo;  
        idxs = sum(CD(:,1:3) == CR,2)==3;
        positions_on_rev = find(idxs);
        if length(r_col_data_2-1+positions_on_rev) > length(OA(1,r_col_data_1:end))
            pregs = unique(P(:,1));
            OA2 = mat2cell(nan(N,r_col_data_1-1+length(positions_on_rev)),ones(1,N),ones(1,r_col_data_1-1+length(positions_on_rev)));
            OA2(:,1:r_col_data_1-1) = OA(:,1:r_col_data_1-1);
            for h = 1:length(pregs)
                OA2(:,r_col_data_1-1+h) = OA(:,r_col_data_1-1+pregs(h));
            end
            OA = OA2;
        end
        Rcop(fl_data:fl_data+N-1,r_col_data_2-1+positions_on_rev) = OA(:,r_col_data_1:end);
        
        %% SCROLL BAR DE PANEL NEGRO SIEMPRE ABAJO
        jScrollPane = findjobj(handles.text4);
        jVSB = jScrollPane.getVerticalScrollBar;
        jVSB.setValue(jVSB.getMaximum);
        drawnow;

    end
    switch COMP
        case 'PCWIN64'
            P2 = [RCToExcelA1(fl_data,r_col_data_2) ':' RCToExcelA1(fl_data+N-1,r_col_data_2+cdL-1)];
            xlswrite([r_path r_file],Rcop(fl_data:fl_data+N-1,r_col_data_2:r_col_data_2+cdL-1),rev_sheet,P2);
            P3 = [RCToExcelA1(fl_data,r_col_data_2+cdL) ':' RCToExcelA1(fl_data+N-1,r_col_data_2+cdL)];
            xlswrite([r_path r_file],stud_data_date(:,r_col_data_1),rev_sheet,P3);
        otherwise
            % CASE OTHERS
            log_s = 'Sistema Operativo no soportado.';
            log_s = fillString(log_s,string_length_console);
            handles.text4.String = [handles.text4.String; log_s];
            return; 
    end
    
    log_s = fillString('Finalizado Correctamente',string_length_console);
    handles.text4.String = [handles.text4.String; log_s];
    jScrollPane = findjobj(handles.text4);
    jVSB = jScrollPane.getVerticalScrollBar;
    jVSB.setValue(jVSB.getMaximum);
    drawnow;
    
catch me_err
    log_me = fillString(['    ' me_err.message],string_length_console);
    handles.text4.String = [handles.text4.String; log_me];
    for er = 1:length(me_err.stack)-1
        [~,name_err,ext_err] = fileparts(me_err.stack(er).file);
        log_err_er = fillString(['    ' name_err ext_err ' | ' me_err.stack(er).name ...
                                 ' | line ' num2str(me_err.stack(er).line) ],string_length_console);
        handles.text4.String = [handles.text4.String; log_err_er];
    end
    drawnow;    
end


function outArr = addResp(outArr,rut,val,j,r_col_rut,r_col_data,handles,string_length_console)
    if length(str2num(str2mat(outArr(:,r_col_rut)))) > 1
        lV = find(~(str2num(str2mat(outArr(:,r_col_rut))) - rut));
    else
        lV = 1;
    end
    if isempty(lV)
        log_s = ['    RUT NO ENCONTRADO : ' num2str(rut)];
        log_s = fillString(log_s,string_length_console);
        handles.text4.String = [handles.text4.String; log_s]; 
    else
        val_check = regexp(num2str(val),'[\0-9]{4}\-[\0-9]{2}\-[\0-9]\s[\0-9]{2}:[\0-9]{2}:[\0-9]{2}');
        if ~isempty(val_check)
            if val_check
                try
                    if isempty(outArr{lV,j+r_col_data-1})
                        curr_date = '1900-01-01 00:00:00';                        
                    else
                        curr_date = outArr{lV,j+r_col_data-1};                        
                    end
                catch date_err
                    curr_date = '1900-01-01 00:00:00';
                end
                %disp(['Comp ' curr_date ' v/s ' val 'last date is ' last_date(curr_date,val)])
                outArr{lV,j+r_col_data-1} = last_date(curr_date,val);
            end
        else
            outArr{lV,j+r_col_data-1} = val; 
        end
    end

function p = isequalCellArray(A,B)
    [fA,cA] = size(A);
    for f = 1:fA
        for c = 1:cA 
            auxA = A{f,c};
            if isequal(auxA,B)
                p(f,c) = 1;
            else
                p(f,c) = 0;
            end
        end
    end
    
function s = fillString(s,L)
    l = length(s);
    if L>l
        fS = char(32*ones(1,L-l));
        s = [s fS];
    else
        s = s(1:L);
    end

function Address = RCToExcelA1(Row, Column)
  % The result can have up to three alpha digits.
  Digits = zeros(1, 3);
  % Convert number-number format to alpha-number format.
  Digits(1) = max(floor(((Column - 1) / 26 - 1) / 26), 0);
  Digits(2) = floor((Column - Digits(1) * 26 * 26 - 1) / 26);
  Digits(3) = rem(Column - 1, 26) + 1;
  % Delete negative numbers and convert blank cells to spaces.
  Digits(Digits > 0) = Digits(Digits > 0) + 64;
  Digits(Digits == 0) = Digits(Digits == 0) + 32;
  % There may be leading spaces, so trim them away.
  Address = strtrim([char(Digits), num2str(Row)]);
  
function date = last_date(date1,date2)
    disp(['DATE 1: ' date1 ' - DATE 2: ' date2])
    date1_1 = strsplit(date1,'-');
    c = strsplit(date1_1{3},' ');
    date1_1 = {date1_1{1:2}, c{:}};
    c = strsplit(date1_1{4},':');
    date1_1 = {date1_1{1:3}, c{:}};
    date1_N = str2double(date1_1);
    
    date2_1 = strsplit(date2,'-');
    c = strsplit(date2_1{3},' ');
    date2_1 = {date2_1{1:2}, c{:}};
    c = strsplit(date2_1{4},':');
    date2_1 = {date2_1{1:3}, c{:}};
    date2_N = str2double(date2_1);

    date_dif = date2_N - date1_N;

    i = find(date_dif,1);
    if isempty(i)
        date = date1;
    else
        if date_dif(i) > 0
            date = date2; 
        else
            date = date1;
        end
    end


        

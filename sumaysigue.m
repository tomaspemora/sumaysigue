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

% Last Modified by GUIDE v2.5 30-Sep-2016 00:37:57

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
warning off; %#ok<*WNOFF>
% Choose default command line output for sumaysigue
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% Icono Suma y Sigue en la ventana
% javaFrame = get(hObject,'JavaFrame');
% javaFrame.setFigureIcon(javax.swing.ImageIcon('icon.jpg'));
warning on; %#ok<*WNON>

% This sets up the initial plot - only do when we are invisible
% so window can get raised using sumaysigue.


% UIWAIT makes sumaysigue wait for user response (see UIRESUME)
% uiwait(handles.figure1);

% --- Outputs from this function are returned to the command line.
function varargout = sumaysigue_OutputFcn(hObject, eventdata, handles) %#ok<*INUSL>
varargout{1} = handles.output;

function CloseMenuItem_Callback(hObject, eventdata, handles) %#ok<*DEFNU>
selection = questdlg(['Close ' get(handles.figure1,'Name') '?'],...
                     ['Close ' get(handles.figure1,'Name') '...'],...
                     'Yes','No','Yes');
if strcmp(selection,'No')
    return;
end
delete(handles.figure1)

function Ayuda_Callback(hObject, eventdata, handles) %#ok<*INUSD>

function Acerca_Callback(hObject, eventdata, handles)

function pushbutton2_Callback(hObject, eventdata, handles)
[filename ,pathname, ~] = uigetfile({'*.xls;*.xlsx;','Archivos Excel'}, 'Seleccione el archivo de plantilla');
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
[filename ,pathname, ~] = uigetfile({'*.xls;*.xlsx;','Archivos Excel'}, 'Seleccione el o los archivos con las respuestas', 'MultiSelect', 'on');
if ~iscell(filename) & filename == 0 %#ok<AND2>
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
    
    string_length_console = 97;                                            % LARGO DE STRING EN CONSOLA
    COMP = computer;
    
    %% EXTRACCION DE ARCHIVOS DE ENTRADA DE USUARIO
    r_file          = handles.pushbutton2.UserData.files;                   % ARCHIVO DE PLANTILLA DE REIVISION
    r_path          = handles.pushbutton2.UserData.path;                    % RUTA DE ARCHIVO DE PLANTILLA DE REVISION
    a_file          = cellstr(handles.pushbutton4.UserData.files);          % ARCHIVO(S) DE RESPUESTAS
    a_path          = handles.pushbutton4.UserData.path;                    % RUTA DE ARCHIVO(S) DE RESPUESTAS
    
    %% EXTRACCION DE PARAMETROS DE CONFIGURACION       
    parametros      = load([pwd filesep 'resources' filesep 'datos.mat']);                      % ARCHIVO DE DATOS DE CONFIGURACION
    parametros      = parametros.f;
    N               = str2double(parametros.data.N);                   % NUMERO DE ALUMNOS
    
    r_fl_data       = str2double(parametros.data.r_fl_data);           % FILA EN QUE EMPIEZA LA DATA DE PLANTILLA
    r_ll_data       = N + r_fl_data - 1;                               % FILA EN QUE TERMINAR LA DATA DE PLANTILLA
    r_col_rut       = str2double(parametros.data.r_col_rut);           % INDICE COLUMNA DE RUT EN PLANTILLA
    r_col_data_1    = str2double(parametros.data.r_col_data_1);        % INDICE COLUMNA EN QUE TERMINAN LOS DATOS DE PERSONAS EN PLANTILLA
    r_col_data_2    = str2double(parametros.data.r_col_data_2);        % INDICE COLUMNA EN QUE EMPIEZAN LAS RESPUESTAS EN PLANTILLA
        
	o_fl_data       = str2double(parametros.data.o_fl_data);           % FILA EN QUE EMPIEZA LA DATA DE PLANTILLA DE PREGUNTAS ABIERTAS
    o_ll_data       = N + o_fl_data - 1;                               % FILA EN QUE TERMINAR LA DATA DE PLANTILLA DE PREGUNTAS ABIERTAS
    o_col_rut       = str2double(parametros.data.o_col_rut);           % INDICE COLUMNA DE RUT EN PLANTILLA DE PREGUNTAS ABIERTAS
    o_col_data_1    = str2double(parametros.data.o_col_data_1);        % INDICE COLUMNA EN QUE TERMINAN LOS DATOS DE PERSONAS EN PLANTILLA DE PREGUNTAS ABIERTAS
    o_col_data_2    = str2double(parametros.data.o_col_data_2);        % INDICE COLUMNA EN QUE EMPIEZAN LAS RESPUESTAS EN PLANTILLA DE PREGUNTAS ABIERTAS
    
    a_col_fecha     = 3;                                               % (Archivos Columna Fecha) INDICE COLUMNA DE FECHA EN ARCHIVOS RESPUESTAS
    a_col_rut       = str2double(parametros.data.a_col_rut);           % (Archivos Columna RUT) INDICE COLUMNA DE RUT EN ARCHIVOS RESPUESTAS
    a_fl_data       = str2double(parametros.data.a_fl_data);           % (Archivos First Line (column) DATA) COLUMNA EN QUE EMPIEZAN LAS RESPUESTAS EN ARCHIVOS DE RESPUESTAS
    info_fil        = 1;                                               % (Fila con parametros del archivo) FILA CON INFO DE ARCHIVOS RESPUESTAS
    
    rev_sheet       = parametros.data.pes_rev;                         % PESTAÑA DE PLANTILLA DE REVISION
    abi_sheet       = parametros.data.pes_abi;                         % PESTAÑA DE PLANTILLA DE PREGUNTAS ABIERTAS
    string_flag     = parametros.data.string_flag;                     % STRING PARA IDENTIFICAR LAS PREGUNTAS  
    
    %<dlf>
    ruts_ignore     = parametros.data.ruts_ignore;                     % RUTs que serán ignorados %dlf
    ruts_ignore     = str2double(strsplit(ruts_ignore,{',',';',' '})); %dlf
    if isnan(ruts_ignore)
        ruts_ignore=[];
        MSGtoConsole('RUTs a ignorar deben estar separados por espacio, no contener guión ni letras',string_length_console,handles);
    end
    %</dlf>     
    
    switch COMP
        case 'MACI64'
        % CASE MAC
            SheetsStruct    = importdata([r_path r_file]);
            Sheets          = fieldnames(SheetsStruct.textdata);
        case {'PCWIN64','PCWIN'}
        % CASE WINDOWS
            [~,Sheets,~]    = xlsfinfo([r_path r_file]);                            % SHEETS: CELL ARRAY CON NOMBRE PESTAÑA EN PLANTILLA (CAMBIAR POR ENTRADA USUARIO)
            Sheets          = cellstr(Sheets);            
        otherwise
        % CASE OTHERS
            MSGtoConsole('Sistema Operativo no soportado.',string_length_console,handles);
            return;            
    end
    % Sheets: Nombres de las pestañas
    
    %% COMPROBAR QUE PESTAÑAS DE REVISION Y PREGUNTAS ABIERTAS EXISTAN EN ARCHIVO DE REVISION
    bp_rev = 0; bp_abi = 0;
    for sh = 1:length(Sheets)
        bp_rev = strcmp(Sheets{sh},rev_sheet) + bp_rev;
        bp_abi = strcmp(Sheets{sh},abi_sheet) + bp_abi;
    end
    if ~bp_rev
        MSGtoConsole(['Revise el nombre de la pestaña de revisión. La pestaña ''' rev_sheet ''' no se encuentra en ''' r_file ''''],string_length_console,handles);
        return;
    elseif ~bp_abi
        MSGtoConsole(['Revise el nombre de la pestaña de preguntas abiertas. La pestaña ''' abi_sheet ''' no se encuentra en ''' r_file ''''],string_length_console,handles);
        return;
    end
    
    %% ABRIR ARCHIVOS DE REVISION
    switch COMP
        case {'PCWIN64','PCWIN'}
        % CASE PCWIN64
            [~,~,R]      = xlsread([r_path r_file],rev_sheet);                   % R: CELL ARRAY CON DATOS PLANTILLA REVISION
            [~,~,O]      = xlsread([r_path r_file],abi_sheet);                   % A: CELL ARRAY CON DATOS PLANTILLA PREGUNTAS ABIERTAS
        otherwise
        % CASE OTHERS
            MSGtoConsole('Sistema Operativo no soportado.',string_length_console,handles);
            return;
    end
    % R y O: Cell array con toda la info de revision y preguntas abiertas respectivamente
    
    stud_data_rev       = R( r_fl_data:r_ll_data , 1:r_col_data_1 - 1 );        % INFORMACION DE LOS ALUMNOS, NO DEBE TENER MAS DE 4 COLUMNAS, i.e r_col_data_1 = 5 siempre    
    course_data_rev     = R( 1:r_fl_data-1 , r_col_data_2:end );                % INDICE DE LOS TALLERES y ACTIVIDADES DEL CURSO
    stud_data_abi       = O( o_fl_data:o_ll_data , 1:o_col_data_1 - 1 );        % INFORMACION (PREGUNTAS ABIERTAS) DE LOS ALUMNOS, NO DEBE TENER MAS DE 4 COLUMNAS, i.e r_col_data_1 = 5 siempre 
    course_data_abi     = O( 1:o_fl_data-2 , o_col_data_2:end );                % INDICE DE LOS TALLERES y ACTIVIDADES DEL CURSO (PREGUNTAS ABIERTAS)
    
    [~, cdL]            = size(course_data_rev);                                % NUMERO DE TALLERES Y ACTIVIDADES    
    [~, cdLo]           = size(course_data_abi);                                % NUMERO DE TALLERES Y ACTIVIDADES    
    
    for i = 1:N
        stud_data_rev{i,r_col_rut} = num2str(stud_data_rev{i,r_col_rut});
        stud_data_abi{i,r_col_rut} = num2str(stud_data_abi{i,r_col_rut});
    end
    stud_data_rev(:,r_col_rut) = multiStrrep(stud_data_rev(:,r_col_rut),{'-','K','k','.',' '},{'','0','0','',''});   % REPARACION DE CASILLA DE RUT (-K,-k -> -0 y borrar .)
    rut_list_rev = str2num(str2mat(stud_data_rev(:,r_col_rut)));
    if isempty(rut_list_rev)
        MSGtoConsole('REVISE LA LISTA DE RUTS DE LA PLANILLA DE REVISION (SOLO 0-9, "-", "k", "K", ".", " ")',string_length_console,handles);
        return;
    end  
    stud_data_abi(:,r_col_rut) = multiStrrep(stud_data_abi(:,r_col_rut),{'-','K','k','.',' '},{'','0','0','',''});   % REPARACION DE CASILLA DE RUT (-K,-k -> -0 y borrar .)                                                                     % COPIA PARA ARREGLO DE RESPUESTAS ABIERTAS
    rut_list_abi = str2num(str2mat(stud_data_abi(:,o_col_rut)));
    if isempty(rut_list_abi)
        MSGtoConsole('REVISE LA LISTA DE RUTS DE LA PLANILLA DE PREGUNTAS ABIERTAS (SOLO 0-9, "-", "k", "K", ".", " ")',string_length_console,handles);
        return;
    end
    stud_data_date      = stud_data_rev;
    % stud_data_rev
    % course_data_rev
    % stud_data_date
    
    %% ARREGLO DE INDICE DE TALLERES Y ACTIVIDADES
    fin_check = [1 1];
    for n = 1:cdL
        if strcmp(course_data_rev{1,n},'fin') || strcmp(course_data_rev{2,n},'fin') || strcmp(course_data_rev{3,n},'fin') || strcmp(course_data_rev{4,n},'fin')
            cdL = n-1;
            fin_check(1) = 0;
            break;
        end
    end
    for n = 1:cdLo
        if strcmp(course_data_abi{1,n},'fin') || strcmp(course_data_abi{2,n},'fin') || strcmp(course_data_abi{3,n},'fin') || strcmp(course_data_abi{4,n},'fin')
            cdLo = n-1;
            fin_check(2) = 0;
            break;
        end
    end
    if fin_check(1)
        MSGtoConsole('Revise que tenga la palabra ''fin'' después de la última pregunta en la planilla de revisión.',string_length_console,handles);return;
    elseif fin_check(2)
        MSGtoConsole('Revise que tenga la palabra ''fin'' después de la última pregunta en la planilla de preguntas abiertas.',string_length_console,handles);return;
    end
    course_data_rev = course_data_rev(:,1:cdL);
    course_data_abi = course_data_abi(:,1:cdLo);
    
    for n = 1:cdL    
        if ischar(course_data_rev{3,n})
            if isempty(str2num(course_data_rev{3,n}))
                MSGtoConsole('Revise los campos de numeración de ''Página'' en la planilla de revisión (0-9).',string_length_console,handles);return;
            else
                course_data_rev{3,n} = str2num(course_data_rev{3,n});
            end
        end
        if isnan(course_data_rev{1,n}), course_data_rev{1,n} = course_data_rev{1,n-1}; end  % RELLENO DE CAMPOS VACIOS EN INDICE DE TALLERES
        if isnan(course_data_rev{2,n}), course_data_rev{2,n} = course_data_rev{2,n-1}; end  % RELLENO DE CAMPOS VACIOS EN INDICE DE ACTIVIDADES
        if isnan(course_data_rev{3,n}), course_data_rev{3,n} = course_data_rev{3,n-1}; end  % RELLENO DE CAMPOS VACIOS EN INDICE DE PAGINAS
    end
    for n = 1:cdLo
        if ischar(course_data_abi{3,n})
            if isempty(str2num(course_data_abi{3,n}))
                MSGtoConsole('Revise los campos de numeración de ''Página'' en la planilla de preguntas abiertas (0-9).',string_length_console,handles);return;
            else
                course_data_abi{3,n} = str2num(course_data_abi{3,n});
            end
        end
        if ischar(course_data_abi{4,n})
            if isempty(str2num(course_data_abi{4,n}))
                MSGtoConsole('Revise los campos de numeración de ''Pregunta'' en la planilla de preguntas abiertas (0-9).',string_length_console,handles);return;
            else
                course_data_abi{4,n} = str2num(course_data_abi{4,n});
            end
        end
        if isnan(course_data_abi{1,n}), course_data_abi{1,n} = course_data_abi{1,n-1}; end  % RELLENO DE CAMPOS VACIOS EN INDICE DE TALLERES
        if isnan(course_data_abi{2,n}), course_data_abi{2,n} = course_data_abi{2,n-1}; end  % RELLENO DE CAMPOS VACIOS EN INDICE DE ACTIVIDADES
        if isnan(course_data_abi{3,n}), course_data_abi{3,n} = course_data_abi{3,n-1}; end  % RELLENO DE CAMPOS VACIOS EN INDICE DE ACTIVIDADES
    end
    
    %% CONVERSION DE INDICE DE TALLERES Y ACTIVIDADES A INDICE NUMERICO
    aux_m = course_data_rev(1:2,1)'; 
    cont  = [1 1]; 
    for n = 1:cdL
        aux = [course_data_rev(1:4,n); 0]';
        if ~strcmp(aux{1},aux_m{1})                                          
            cont(1) = cont(1) + 1; cont(2) = 1;
        else
            if ~strcmp(aux{2},aux_m{2}), cont(2) = cont(2) + 1; end
        end
        if strcmp(aux{4},'-'), aux{4} = -1;                                               % VERIFICANDO TIPO DE PREGUNTA ABIERTA O NO ABIERTA
        else
            if ischar(aux{4})                                               % CASO PREGUNTA ABIERTA O TIPO TEXTO ('4a', '3b', etc)
                aux_copy = aux{4};
                saux = strsplit(aux{4},'(A)');
                if length(saux) > 1, aux{5} = 1; end                        % CASO PREGUNTA ABIERTA
                aux{4} = str2double(saux{1});
                aux_copy = regexprep(regexprep(aux_copy,'[^\0-9]',''),'[\( \)]','');        
                if isnan(aux{4}) | (real(aux{4}) ~= aux{4}), aux{4} = str2double(aux_copy); end %#ok<OR2>
            end
        end
        CD(n,:) = [cont aux{3:5}];                                          % INDICE DE ACTIVIDADES [TALLER | ACTIVIDAD | PAGINA | PREGUNTA | ABIERTA/NOABIERTA]
        aux_m = aux(1:2);
    end
    % CD: INDICE NUMERICO DE LA PLANILLA DE REVISION
    aux_m = course_data_abi(1:2,1)'; 
    cont  = [1 1];
    for n = 1:cdLo
        aux = course_data_abi(1:4,n)';
        if ~strcmp(aux{1},aux_m{1})                                          
            cont(1) = cont(1) + 1; cont(2) = 1;
        else
            if ~strcmp(aux{2},aux_m{2}), cont(2) = cont(2) + 1; end
        end
        CDO(n,:) = [cont aux{3:4}];                                          % #ok<NASGU> INDICE DE ACTIVIDADES [TALLER | ACTIVIDAD | PAGINA | PREGUNTA | ABIERTA/NOABIERTA]
        aux_m = aux(1:2);
    end
    clearvars cont aux aux_m saux aux_copy n sh bp_rev bp_abi fin_check i
    % RESTRICCIONES
    % TALLER: SOLO SE ADMITE TEXTO PARA LAS CASILLAS, ALFANUMERICO
    % ACTIVIDAD: SOLO SE ADMITE TEXTO PARA LAS CASILLAS ALFANUMERICO
    % PAGINA: SOLO SE ADMITE UN NUMERO PARA LAS CASILLAS
    % PREGUNTA: ADMITE NUMERO O NUMERO Y LETRAS O "-": ej: 1, 2, 3a , 3b, 4(A), -

    %% REVISANDO TODOS LOS ARCHIVOS DE RESPUESTAS SELECCIONADOS
    for z = 1:length(a_file)
        
        MSGtoConsole(['Revisando el archivo ' a_file{z} ' ...'],string_length_console,handles);  
        clear P OA
        copia_data_rev          = stud_data_rev;                                % VARIABLE DE COPIA DE INFORMACION DE LOS ALUMNOS  
        copia_data_abi          = stud_data_abi;                                % VARIABLE DE COPIA DE INFORMACION DE LOS ALUMNOS (Preguntas Abiertas) 
        codigo                  = a_file{z};                                    % NOMBRE CODIGO DE ARCHIVO DE RESPUESTAS ej: "variables TXAY_ZZ.xls"
        idx                     = regexp(codigo,'t[\0-9]a[\0-9][_\0-9]');      	% INDICE DE DONDE APARECE TXAY_ZZ.xl...
        codigo                  = strrep(regexprep(codigo(idx:end),'\W\w{1,10}',''),'_','');  % CODIGO LIMPIO TXAYZZ
        switch COMP
            case {'PCWIN64','PCWIN'}
                [~,~,A]         = xlsread([a_path a_file{z}]);                  % CELL ARRAY CON INFO DE ARCHIVO DE RESPUESTAS
            otherwise
                % CASE OTHERS
                MSGtoConsole('Sistema Operativo no soportado.',string_length_console,handles);
                return;
        end        
        info                        = A(info_fil,a_fl_data:end);                                                           % PRIMERA FILA DE ARCHIVO DE RESPUESTAS. CONTIENE SUS PARAMETROS
        data                        = A((info_fil+1):end,:);                                                               % DATA DE ARCHIVO DE RESPUESTSAS
        I                           = length(info);                                                                        % NUMERO COLUMNAS ARCHIVO DE RESPUESTAS
        [D,~]                       = size(data);                                                                          % NUMERO DE ENTRADAS DE ARCHIVO DE RESPUESTAS (FILAS)
        for i = 1:I                                                                                                        % CONSTRUCCION DE INDICE DE PREGUNTAS PRESENTES EN ARCHIVOS DE RESPUESTAS
            if isempty(strfind(info{i},string_flag))
                MSGtoConsole(['    Columna extraña encontrada en ' a_file{z} ': col: ' num2str(i+a_fl_data-1) ' - "' info{i} '"'],string_length_console,handles);
                continue;
            end
            P(i,:)  	= str2num(multiStrrep(info{i},{string_flag,'_'},{'',' '}));                                        % P CONTIENE LOS HEADERS DEL ARCHIVOS DE RESPUESTAS (HTML) 1º COL (PREG) 2º COL (SUBPREG)
        end
        [pL,~]          = size(P);                                                                                         % NUMERO DE SUBPREGUNTAS EN ARCHIVO RESPUESTAS 
        
        %% EXTRACCION Y SINTESIS DE RESPUESTAS DESDE ARCHIVO DE RESPUESTAS (FOR SOBRE ENTRADAS (FILAS) DE ARCHIVO DE RESPUESTAS)
        jo = [];io = 1;
        clear data_o
        for i = 1:D            
            try
                rut = str2num(multiStrrep(data{i,a_col_rut},{'-','K','k','.',' '},{'','0','0','',''})); %#ok<*ST2NM>
            catch
                MSGtoConsole(['REVISE LOS RUTS DEL ARCHIVO ' a_file{z} ' (SOLO 0-9, "-", "k", "K", ".", " ")'],string_length_console,handles);
                break;
            end
            fecha   = data{i,a_col_fecha};
            for j = a_fl_data:(pL + a_fl_data - 1)  % FOR SOBRE PREGUNTAS DE ARCHIVO DE RESPUESTAS (COLUMNAS)
                resp = data{i,j};
                
                if ~isempty(resp)
                    resp = strsplit(resp,'#');
                    ab_resp = resp{3};
                    resp = strsplit(resp{1}, ':');
                    s    = resp{1};
                    
                    switch s
                        case 'buena'    , val = 1;
                        case 'mala'     , val = 0;
                        case 'intento'  , val = NaN;
                        case 'enviada'  , val = 2; jo = unique([jo j]); data_o{io,1}=ab_resp; data_o{io,2}=rut;data_o{io,3} = j;io=io+1;                        
                    end
                    
                    if ~isnan(val)
                        copia_data_rev = addResp(copia_data_rev,rut,val,            j - a_fl_data + 1 , r_col_rut , r_col_data_1,   1,handles,string_length_console,ruts_ignore); % dlf
                    end
                    
                    break;
                end
            end
            stud_data_date = addResp(stud_data_date,rut,fecha,a_fl_data+1,r_col_rut,0,0,handles,string_length_console,ruts_ignore); % dlf
        end
        if ~isempty(jo)
            for i = 1:length(data_o)            
                copia_data_abi = addResp(copia_data_abi,data_o{i,2},data_o{i,1}, find(data_o{i,3} == jo) , o_col_rut , o_col_data_1,   0,handles,string_length_console,ruts_ignore); % dlf
            end
        end
        data_rev_final = stud_data_rev;
        
        %% POST PROCESS
        NP = max(P(:,1));
        for i = 1:NP
            clear OAC
            G = find(~(P(:,1)-i));                                      % columnas asociadas a subpreguntas de la misma pregunta (puede ser mas de una)
            try                                                         % Tratamos de ordenarlas, si es que existen
                aux = copia_data_rev(:,4+(G(1):G(end)));                % Recuperamos las respuestas
                [faux,caux] = size(aux);                                    
                Iaux = isequalCellArray(aux,[]);                       
                for j = 1:faux
                   for g = 1:caux
                       if Iaux(j,g)
                           aux{j,g} = NaN;
                       end
                   end
                end
                OAC = cell2mat(aux);                                    % Preparamos el arreglo para poder colapsarlo a una sola columna (en caso de tener mas de una) 
                if length(G) > 1
                    OAC = prod(OAC,2);                                  % Colapsamos la columna con prod
                end
                data_rev_final(:,i+4) = mat2cell(OAC,ones(1,length(OAC)),1); %#ok<MMTC> % Copiamos la columna a la columna correspondiente i
            catch
                MSGtoConsole(['    La pregunta ' num2str(i) ' no se respondía o nadie la ha respondido aún!'],string_length_console,handles);
            end  
        end
        
        %% ORDENAMIENTO DEL HTML ACTUAL EN EL ARREGLO FINAL
        codigo_page = codigo(end-1:end);
        codigo(end-1:end) = [];
        c_codigo = str2double(strsplit(codigo,'[/t/a]','DelimiterType','RegularExpression'));
        c_codigo(1) = [];
        c_codigo(3) = str2double(codigo_page);
        
        idxs = sum(CD(:,1:3) == (ones(cdL,1) * c_codigo),2) == 3;
        idxsO = sum(CDO(:,1:3) == (ones(cdLo,1) * c_codigo),2) == 3;
        positions_on_abi = find(idxsO);
        positions_on_rev = find(idxs);
        
        if length(data_rev_final(1,r_col_data_1:end)) > 0 %#ok<ISMT>                                  % Revisamos si hay algo que agregar al arreglo final
            if length(positions_on_rev) > length(data_rev_final(1,r_col_data_1:end))                  % DEJANDO COLUMNA EN BLANCO PARA PREGUNTA QUE NO APARECIO EN HTML
                pregs = unique(P(:,1));
                OA2 = mat2cell(nan(N,r_col_data_1-1+length(positions_on_rev)),ones(1,N),ones(1,r_col_data_1-1+length(positions_on_rev)));       %#ok<MMTC>
                OA2(:,1:r_col_data_1-1) = data_rev_final(:,1:r_col_data_1-1);
                for h = 1:length(pregs)
                    OA2(:,r_col_data_1-1+h) = data_rev_final(:,r_col_data_1-1+pregs(h));
                end
                data_rev_final = OA2;
            end
            try
                R(r_fl_data:r_fl_data+N-1,r_col_data_2-1+positions_on_rev) = data_rev_final(:,r_col_data_1:end);
            catch
                MSGtoConsole(['Ocurrio un error - Podría ser que tenga mal indexado el archivo ' a_file{z} ' en la planilla'],string_length_console,handles);
            end
        end
        if ~isempty(copia_data_abi(1,o_col_data_1:end))
            if length(positions_on_abi) > length(copia_data_abi(1,o_col_data_1:end))                  % DEJANDO COLUMNA EN BLANCO PARA PREGUNTA QUE NO APARECIO EN HTML
                OA3 = mat2cell(nan(N,o_col_data_1-1+length(positions_on_abi)),ones(1,N),ones(1,o_col_data_1-1+length(positions_on_abi)));       %#ok<MMTC>
                OA3(:,1:o_col_data_1-1) = copia_data_abi(:,1:o_col_data_1-1);
                for h = 1:length(jo)
                    OA3(:,o_col_data_1-1+h) = copia_data_abi(:,o_col_data_1-1+jo(h));
                end
                copia_data_abi = OA3;
            end 
            try
                O(o_fl_data:o_fl_data+N-1,o_col_data_2-1+positions_on_abi) = copia_data_abi(:,o_col_data_1:end);
            catch
                MSGtoConsole(['Ocurrio un error - Podría ser que tenga mal indexado el archivo ' a_file{z} ' en la planilla'],string_length_console,handles);
            end
        end
        %% SCROLL BAR DE PANEL NEGRO SIEMPRE ABAJO
        jScrollPane = findjobj(handles.text4);jVSB = jScrollPane.getVerticalScrollBar;jVSB.setValue(jVSB.getMaximum);
        drawnow;

    end
    switch COMP
        case {'PCWIN64','PCWIN'}
            P2 = [RCToExcelA1(r_fl_data,r_col_data_2) ':' RCToExcelA1(r_fl_data+N-1,r_col_data_2+cdL-1)];
            xlswrite([r_path r_file],R(r_fl_data:r_fl_data+N-1,r_col_data_2:r_col_data_2+cdL-1),rev_sheet,P2);
            
            P3 = [RCToExcelA1(r_fl_data,r_col_data_2+cdL) ':' RCToExcelA1(r_fl_data+N-1,r_col_data_2+cdL)];
            xlswrite([r_path r_file],stud_data_date(:,r_col_data_1),rev_sheet,P3);
            
            P4 = [RCToExcelA1(o_fl_data,o_col_data_2) ':' RCToExcelA1(o_fl_data+N-1,o_col_data_2+cdLo-1)];
            xlswrite([r_path r_file],O(o_fl_data:o_fl_data+N-1,o_col_data_2:o_col_data_2+cdLo-1),abi_sheet,P4);
        otherwise
            % CASE OTHERS
            MSGtoConsole('Sistema Operativo no soportado.',string_length_console,handles);
            return; 
    end
    MSGtoConsole('Finalizado Correctamente',string_length_console,handles);
    jScrollPane = findjobj(handles.text4);jVSB = jScrollPane.getVerticalScrollBar;jVSB.setValue(jVSB.getMaximum);
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


function outArr = addResp(outArr,rut,val, j , col_rut , col_data , log_disp,handles,string_length_console,ruts_ignore) % dlf
    if length(str2num(str2mat(outArr(:,col_rut)))) > 1 %#ok<*DSTRMT>
        rut_list = str2num(str2mat(outArr(:,col_rut)));
        rut_list(isnan(rut_list)) = -1;
        rut_check = rut_list - rut;
        lV = find(~rut_check);
        if length(lV) > 1
            log_s = ['    RUTS DUPLICADOS EN PLANILLA: Nº ' num2str(lV(1)) ' y Nº ' num2str(lV(2))];
            log_s = fillString(log_s,string_length_console);
            handles.text4.String = [handles.text4.String; log_s];
            lV = lV(1);
        end
    else
        lV = 1;
    end
    if isempty(lV) 
        rut_check = ruts_ignore - rut; % dlf
        lVi = find(~rut_check); % dlf
        %if log_disp && rut ~= 235772043 && rut ~= 300001017 && rut ~= 70179555 % dlf 
        if log_disp && isempty(lVi) % dlf
            log_s = ['    RUT NO ENCONTRADO : ' num2str(rut)];
            log_s = fillString(log_s,string_length_console);
            handles.text4.String = [handles.text4.String; log_s]; 
        end
    else
        val_check = regexp(num2str(val),'[\0-9]{4}\-[\0-9]{2}\-[\0-9]\s[\0-9]{2}:[\0-9]{2}:[\0-9]{2}');
        if ~isempty(val_check)
            if val_check
                try
                    if isempty(outArr{lV,j + col_data - 1})
                        curr_date = '1900-01-01 00:00:00';                        
                    else
                        curr_date = outArr{lV,j + col_data-1};                        
                    end
                catch
                    curr_date = '1900-01-01 00:00:00';
                end
                outArr{lV,j + col_data - 1} = last_date(curr_date,val);
            end
        else
            outArr{lV,j + col_data - 1} = val;
        end
    end

function p = isequalCellArray(A,B)
    [fA,cA] = size(A);
    for f = 1:fA
        for c = 1:cA 
            auxA = A{f,c};
            if isequal(auxA,B)
                p(f,c) = 1; %#ok<*AGROW>
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
    date1_1 = {date1_1{1:2}, c{:}}; %#ok<*CCAT>
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


function MSGtoConsole(string,string_length_console,handles)    
    log_s = string;
    log_s = fillString(log_s,string_length_console);
    handles.text4.String = [handles.text4.String; log_s];
    
function string = multiStrrep(string,cell1,cell2)
    cell1 = cellstr(cell1);
    cell2 = cellstr(cell2);
    if (numel(cell1) == length(cell1)) && (numel(cell2) == length(cell2))
        if numel(cell1) == numel(cell2)            
            for msr = 1:numel(cell1)
                string = strrep(string,cell1{msr},cell2{msr});
            end
        else
            disp('Ambos Cell Array deben tener el mismo largo');
        end
    else
        disp('Ambos Cell Array deben ser vectores');
    end

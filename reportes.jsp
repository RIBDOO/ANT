<%-- 
    Document   : Reporte
    Created on : 12/11/2016, 03:50:14 PM
    Author     : User
    Project    : Reportes
--%>

<%@page import="impresiones.Orden_Serv"%>
<%@page import="impresiones.Mov_OrdenT"%>
<%@page import="impresiones.Orden_Trab"%>
<%@page import="impresiones.Mov_Orden"%>
<%@page import="impresiones.GuiaM"%>
<%@page import="impresiones.Orden"%>
<%@page import="net.sf.jasperreports.engine.export.JRXlsExporter"%>
<%@page import="net.sf.jasperreports.engine.export.JRXlsExporterParameter"%>
<%@page import="impresiones.Lectura_Archivo"%>
<%@page import="net.sf.jasperreports.engine.util.JRLoader"%>
<%@page import="net.sf.jasperreports.engine.data.JRBeanCollectionDataSource"%>
<%@page import="impresiones.Conexion"%>
<%@page import="java.text.DecimalFormat"%>
<%@page import="java.text.NumberFormat"%>
<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@page import="net.sf.jasperreports.engine.*" %>
<%@page import="java.util.*"%>
<%@page import="java.io.*"%>
<%@page import="java.sql.*"%>
<%
	String n_arc = request.getParameter("n_Arc"); 
        String t_imp = request.getParameter("t_Imp");
        String nro_doc = request.getParameter("nroDoc");
        String t_sal = request.getParameter("t_Sal");
        // PARA REPOSICION DE PRODUCTOS
        String parAdi = request.getParameter("parAdi");
        String addCond = request.getParameter("addCond");
        
        
        
        String cNAME = request.getParameter("parNAME");
        String cRUC  = request.getParameter("parRUC");
        String cDIRE = request.getParameter("parDIRE");
        String cTELF = request.getParameter("parTELEF");
        
        String pTIT = request.getParameter("pTITULO");
        
        String formato = "", formato_ex = "";
        File reportFile;
        byte[] bytes = null;
        
        /* SE OBTIENE (NOMBRE BD, USER, PASS) */
        FileReader f = new FileReader(application.getRealPath("/conexion/"+n_arc+".txt"));
        Lectura_Archivo lector = new Lectura_Archivo(f);
        lector.lee_archivo();
        
        JasperPrint print = null;
        JasperPrint jasperPrint = null; 
        Map parametros;
        
        Random rd = new Random();
        
        /* SE CREA CONEXION A LA BD */
        Connection conn = null;
        try {
            String url = "jdbc:db2://"+lector.get_DB();
            Class.forName("com.ibm.db2.jcc.DB2Driver");
            conn = DriverManager.getConnection(url,lector.get_User(),lector.get_Pass());
        }
        catch (SQLException e) {
            System.out.println("Error de conexión: " + e.getMessage());
            System.exit(4);
        }
        
        /* Variables para exportar Excel */
        OutputStream ouputStream = null;
        JRXlsExporter exporterXLS;
        ByteArrayOutputStream arrayOutputStream;
        List sheets = new ArrayList();
        
        /*AGREGADO 20181001 para pprueba de ticket*/
        if(t_imp.equals("rep_ticket")){
            
            
        }
        //################################################
        if(t_imp.equals("Fact")){
            formato = "rep_fact";
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            
            parametros = new HashMap();
            parametros.clear();
            parametros.put("NRO_DOC",nro_doc);
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            conn.close();
            /********************PDF********************/
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            print = jasperPrint;
        }
        /*
        if(t_imp.equals("Bolt")){
            formato = "rep_bolt";
        }
        */
        if(t_imp.equals("GuiaM")){
            formato = "rep_guiam";
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            GuiaM guia = new GuiaM(nro_doc,lector.get_DB(),lector.get_User(),lector.get_Pass());
            guia.ejecuta_query();
            /********************PDF********************/
            bytes = JasperRunManager.runReportToPdf(reportFile.getPath(), guia.get_map(), new JRBeanCollectionDataSource(guia.get_list()));
            /********************EXCEL********************/
            print = JasperFillManager.fillReport(reportFile.getPath(),guia.get_map(),new JRBeanCollectionDataSource(guia.get_list()));
        }
        if(t_imp.equals("Ordn")){
            formato = "rep_ordn";
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            Orden orden = new Orden(nro_doc,lector.get_DB(),lector.get_User(),lector.get_Pass());
            orden.ejecuta_query();
            /********************PDF********************/
            bytes = JasperRunManager.runReportToPdf(reportFile.getPath(), orden.get_map(), new JRBeanCollectionDataSource(orden.get_list()));
            /********************EXCEL********************/
            print = JasperFillManager.fillReport(reportFile.getPath(),orden.get_map(),new JRBeanCollectionDataSource(orden.get_list()));  
        }
        if(t_imp.equals("MovO")){
            formato = "rep_mov_ordn";
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            Mov_Orden m_orden = new Mov_Orden(nro_doc,lector.get_DB(),lector.get_User(),lector.get_Pass());
            m_orden.ejecuta_query();
            /********************PDF********************/
            bytes = JasperRunManager.runReportToPdf(reportFile.getPath(), m_orden.get_map(), new JRBeanCollectionDataSource(m_orden.get_list()));
            /********************EXCEL********************/
            print = JasperFillManager.fillReport(reportFile.getPath(),m_orden.get_map(),new JRBeanCollectionDataSource(m_orden.get_list()));  
        }
        if(t_imp.equals("OrdT")){
            formato = "rep_ordt";
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            Orden_Trab ordent = new Orden_Trab(nro_doc,lector.get_DB(),lector.get_User(),lector.get_Pass());
            ordent.ejecuta_query();
            /********************PDF********************/
            bytes = JasperRunManager.runReportToPdf(reportFile.getPath(), ordent.get_map(), new JRBeanCollectionDataSource(ordent.get_list()));
            /********************EXCEL********************/
            print = JasperFillManager.fillReport(reportFile.getPath(),ordent.get_map(),new JRBeanCollectionDataSource(ordent.get_list()));  
        }
        if(t_imp.equals("MovT")){
            formato = "rep_mov_ordt";
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            Mov_OrdenT m_ordent = new Mov_OrdenT(nro_doc,lector.get_DB(),lector.get_User(),lector.get_Pass());
            m_ordent.ejecuta_query();
            /********************PDF********************/
            bytes = JasperRunManager.runReportToPdf(reportFile.getPath(), m_ordent.get_map(), new JRBeanCollectionDataSource(m_ordent.get_list()));
            /********************EXCEL********************/
            print = JasperFillManager.fillReport(reportFile.getPath(),m_ordent.get_map(),new JRBeanCollectionDataSource(m_ordent.get_list()));  
        }
        if(t_imp.equals("OrdS")){
            formato = "rep_ords";
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            Orden_Serv ordens = new Orden_Serv(nro_doc,lector.get_DB(),lector.get_User(),lector.get_Pass());
            ordens.ejecuta_query();
            /********************PDF********************/
            bytes = JasperRunManager.runReportToPdf(reportFile.getPath(), ordens.get_map(), new JRBeanCollectionDataSource(ordens.get_list()));
            /********************EXCEL********************/
            print = JasperFillManager.fillReport(reportFile.getPath(),ordens.get_map(),new JRBeanCollectionDataSource(ordens.get_list()));  
        }
        if(t_imp.equals("rep_dstk")){
            formato = "rep_dstk";
            formato_ex = "rep_dstk_ex";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("IN_PAR",addCond);
            if(parAdi.equals("CDG_PROD")){
                parametros.put("PAR_ADD",parAdi);
            }
            if(parAdi.equals("CDG_EQV")){
                parametros.put("PAR_ADD",parAdi);
            }
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);  
            print = jasperPrint;
            conn.close();
        }        
        if(t_imp.equals("rep_estk")){
            formato = "rep_estk";
            formato_ex = "rep_estk_ex";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("IN_PAR",addCond);
            if(parAdi.equals("CDG_PROD")){
                parametros.put("PAR_ADD",parAdi);
            }
            if(parAdi.equals("CDG_EQV")){
                parametros.put("PAR_ADD",parAdi);
            }
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);  
            print = jasperPrint;
            conn.close();
        }
        if(t_imp.equals("rep_pprv")){
            formato = "rep_pprv";
            formato_ex = "rep_pprv_ex";    
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("IN_PAR",addCond);
            if(parAdi.equals("CDG_PROD")){
                parametros.put("PAR_ADD",parAdi);
            }
            if(parAdi.equals("CDG_EQV")){
                parametros.put("PAR_ADD",parAdi);
            }
                                    
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);  
            print = jasperPrint;
            conn.close();
        }
        if(t_imp.equals("rep_stks")){
            formato = "rep_stks";
            formato_ex = "rep_stks_ex";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar datos de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar datos de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar datos de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar datos de la empresa
            parametros.put("IN_PAR",addCond);
            if(parAdi.equals("CDG_PROD")){
                parametros.put("PAR_ADD",parAdi);
            }
            if(parAdi.equals("CDG_EQV")){
                parametros.put("PAR_ADD",parAdi);
            }
            
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            conn.close();
            /********************PDF********************/
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            print = jasperPrint;
        }
        if(t_imp.equals("rep_stks_prod")){
            formato = "rep_stks_prod";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("IN_PAR",addCond);
            if(parAdi.equals("CDG_PROD")){
                parametros.put("PAR_ADD",parAdi);
            }
            if(parAdi.equals("CDG_EQV")){
                parametros.put("PAR_ADD",parAdi);
            }
            
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            conn.close();
            /********************PDF********************/
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            print = jasperPrint;
        }
        if(t_imp.equals("rep_kard")){
            formato = "rep_kard";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("IN_PAR",addCond);
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar datos de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar datos de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar datos de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar datos de la empresa
            if(parAdi.equals("CDG_PROD")){
                parametros.put("PAR_ADD",parAdi);
            }
            if(parAdi.equals("CDG_EQV")){
                parametros.put("PAR_ADD",parAdi);
            }
            
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            conn.close();
            /********************PDF********************/
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            print = jasperPrint;
        }

        if(t_imp.equals("rep_notas_ingreso2")){
            String desde = request.getParameter("parFDES"); 
            String hasta = request.getParameter("parFHAS"); 
            String area = request.getParameter("parAREA"); 
            
            formato = "rep_notas_ingreso2";
            formato_ex = "rep_notas_ingreso2_ex";                            
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("IN_PAR",addCond);
            parametros.put("FDES",desde);
            parametros.put("FHAS",hasta);
            parametros.put("DAREA",area);
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar datos de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar datos de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar datos de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar datos de la empresa
            if(parAdi.equals("CDG_PROD")){
                parametros.put("PAR_ADD",parAdi);
            }
            if(parAdi.equals("CDG_EQV")){
                parametros.put("PAR_ADD",parAdi);
            }
                                    
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();
        }
        if(t_imp.equals("rep_notas_salida3")){
            String desde = request.getParameter("parFDES"); 
            String hasta = request.getParameter("parFHAS"); 
            String area = request.getParameter("parAREA"); 
            
            formato = "rep_notas_salida3";
            formato_ex = "rep_notas_salida3_ex";                            
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("IN_PAR",addCond);
            parametros.put("FDES",desde);
            parametros.put("FHAS",hasta);
            parametros.put("DAREA",area);
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar datos de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar datos de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar datos de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar datos de la empresa
            
            if(parAdi.equals("CDG_PROD")){
                parametros.put("PAR_ADD",parAdi);
            }
            if(parAdi.equals("CDG_EQV")){
                parametros.put("PAR_ADD",parAdi);
            }                       
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();
        }
        if(t_imp.equals("rep_neg")){
            formato = "rep_neg";
            formato_ex = "rep_neg_ex";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar datos de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar datos de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar datos de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar datos de la empresa
            parametros.put("IN_PAR",addCond);
            if(parAdi.equals("CDG_PROD")){
                parametros.put("PAR_ADD",parAdi);
            }
            if(parAdi.equals("CDG_EQV")){
                parametros.put("PAR_ADD",parAdi);
            }
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();
        }
        if(t_imp.equals("rep_kard2")){
            
            String desde = request.getParameter("parFDES"); 
            String hasta = request.getParameter("parFHAS"); 
            
            formato = "rep_kard2";
            formato_ex = "rep_kard2_ex";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar datos de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar datos de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar datos de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar datos de la empresa
            parametros.put("IN_PAR",addCond);
            parametros.put("FDES",desde);
            parametros.put("FHAS",hasta);
            if(parAdi.equals("CDG_PROD")){
                parametros.put("PAR_ADD",parAdi);
            }
            if(parAdi.equals("CDG_EQV")){
                parametros.put("PAR_ADD",parAdi);
            }
            
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            conn.close();
            /********************PDF********************/
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            print = jasperPrint;
        }
        if(t_imp.equals("rep_kard2_resumido")){
            String desde = request.getParameter("parFDES"); 
            String hasta = request.getParameter("parFHAS"); 
            
            formato = "rep_kard2_resumido";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar datos de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar datos de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar datos de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar datos de la empresa
            parametros.put("IN_PAR",addCond);
            parametros.put("FDES",desde);
            parametros.put("FHAS",hasta);
            if(parAdi.equals("CDG_PROD")){
                parametros.put("PAR_ADD",parAdi);
            }
            if(parAdi.equals("CDG_EQV")){
                parametros.put("PAR_ADD",parAdi);
            }
            
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            conn.close();
            /********************PDF********************/
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            print = jasperPrint;
        }
        if(t_imp.equals("rep_stkd")){
            String rep = request.getParameter("parREP"); 
            
            if(rep.equals("1"))
                formato = "rep_stkd";
            else
                formato = "rep_stkd_ex";                            
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar datos de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar datos de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar datos de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar datos de la empresa
            parametros.put("IN_PAR",addCond);
            if(parAdi.equals("CDG_PROD")){
                parametros.put("PAR_ADD",parAdi);
            }
            if(parAdi.equals("CDG_EQV")){
                parametros.put("PAR_ADD",parAdi);
            }

            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            /********************PDF********************/            
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/          
            print = jasperPrint;
            conn.close();
        }
        if(t_imp.equals("rep_lista")){
            formato = "rep_lista";
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            
            parametros = new HashMap();
            parametros.clear();
            parametros.put("PLISTA",parAdi);
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            conn.close();
            /********************PDF********************/
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            print = jasperPrint;
        }
        if(t_imp.equals("det_prod")){
            String rep = request.getParameter("parREP"); 
            String ref = request.getParameter("parREF");
            formato = "det_prod";
            
            StringTokenizer stoken = new StringTokenizer(ref,"|");
            parametros = new HashMap();
            parametros.clear();
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar RUC de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar NOMBRE de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar DIRECCION de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar TELEFONO de la empresa
            int i = 0;
            
            if(rep.equals("1"))
                formato = "det_prod";
            if(rep.equals("2")){
                formato = "det_prod_ref"; 
                while(stoken.hasMoreElements()){
                    String token = stoken.nextElement().toString();
                    i++;
                    parametros.put("REF"+i,token);
                }
            }
            if(rep.equals("3"))
                formato = "det_prod_ubic";                          
                            
            parametros.put("IN_PAR",addCond);
            
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            /********************PDF********************/            
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/          
            print = jasperPrint;
            conn.close();
        }
        if(t_imp.equals("det_dprod")){
            String rep = request.getParameter("parREP"); 
            String area = request.getParameter("parAREA"); 
            String desc = request.getParameter("parDESP"); 
            String prod = request.getParameter("parPROD"); 
            formato = "det_dprod";
            
            parametros = new HashMap();
            parametros.clear();
            
            if(rep.equals("1"))
                formato = "det_dprod_ubic";
            if(rep.equals("2")){
                formato = "det_dprod"; 
            }                        
                            
            parametros.put("IN_PAR",parAdi);
            parametros.put("AREA",area);
            parametros.put("DES_PROD",desc);
            parametros.put("CDG_PROD",prod); 
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar RUC de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar NOMBRE de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar DIRECCION de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar TELEFONO de la empresa
            
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            /********************PDF********************/            
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/          
            print = jasperPrint;
            conn.close();
        }
        if(t_imp.equals("rep_formula")){                                    
            formato = "rep_formula";
            formato_ex = "rep_formula_ex";                                       
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("IN_PAR",addCond);          
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);            
            print = jasperPrint;
            conn.close();
        }
        if(t_imp.equals("rep_ctec")){
            String desde = request.getParameter("parFDES"); 
            String hasta = request.getParameter("parFHAS"); 
            
            formato = "rep_ctec";
            formato_ex = "rep_ctec_ex";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("dDESDE",desde);
            parametros.put("dHASTA",hasta);
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();            
        }
          if(t_imp.equals("rep_ctep")){
            String desde = request.getParameter("parFDES"); 
            String hasta = request.getParameter("parFHAS"); 
            
            formato = "rep_ctep";
            formato_ex = "rep_ctep";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("dDESDE",desde);
            parametros.put("dHASTA",hasta);
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar RUC de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar NOMBRE de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar DIRECCION de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar TELEFONO de la empresa
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();            
        }
           if(t_imp.equals("reg_compras_sunat")){
         //   String desde = request.getParameter("parFDES"); 
         //   String hasta = request.getParameter("parFHAS"); 
            
            formato = "reg_compras_sunat";
            formato_ex = "reg_compras_sunat";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar RUC de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar NOMBRE de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar DIRECCION de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar TELEFONO de la empresa
          //  parametros.put("dDESDE",desde);
          //  parametros.put("dHASTA",hasta);
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();            
        }
              if(t_imp.equals("rep_fecha_ent")){
         //   String desde = request.getParameter("parFDES"); 
         //   String hasta = request.getParameter("parFHAS"); 
            
            formato = "rep_fecha_ent";
            formato_ex = "rep_fecha_ent_ex";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
          //  parametros.put("dDESDE",desde);
          //  parametros.put("dHASTA",hasta);
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();            
        }
              if(t_imp.equals("rep_valp")){
         //   String desde = request.getParameter("parFDES"); 
         //   String hasta = request.getParameter("parFHAS"); 
            
            formato = "rep_valp";
            formato_ex = "rep_valp";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
          //  parametros.put("dDESDE",desde);
          //  parametros.put("dHASTA",hasta);
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();            
        }
        if(t_imp.equals("rep_reso")){
            String desde = request.getParameter("parFDES"); 
            String hasta = request.getParameter("parFHAS");
           
            
            formato = "rep_reso";
            formato_ex = "rep_reso";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("dDESDE",desde);
            parametros.put("dHASTA",hasta);
            parametros.put("nroDoc",nro_doc);
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar datos de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar datos de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar datos de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar datos de la empresa
                     
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();            
        }
        
       //reporte de productos 
       if(t_imp.equals("rep_prod")){
            //String desde = request.getParameter("parFDES"); 
            //String hasta = request.getParameter("parFHAS"); 
            
            formato = "rep_prod";
            formato_ex = "rep_prod";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            //parametros.put("dDESDE",desde);
            //parametros.put("dHASTA",hasta);
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar datos de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar datos de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar datos de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar datos de la empresa
                     
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();            
        }
           //reporte de productos 
       if(t_imp.equals("rep_set")){
            //String desde = request.getParameter("parFDES"); 
            //String hasta = request.getParameter("parFHAS"); 
            
            formato = "rep_set";
            formato_ex = "rep_set";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            //parametros.put("dDESDE",desde);
            //parametros.put("dHASTA",hasta);
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar datos de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar datos de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar datos de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar datos de la empresa
            parametros.put("CDG_OPC",cTELF);      //agregado para elegir entre codigo ANT o EQV
            
                     
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();            
        }           
        if(t_imp.equals("rep_anvenc")){ 
            int index;
                        
            for(index=1; index<=3; index++){                  
                formato = "rep_anvenc" + index;

                reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));

                parametros = new HashMap();
                parametros.clear();

                /********************EXCEL********************/
                reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
                print = jasperPrint;
                sheets.add(print);                               
            }
            conn.close();         
        }
        
       
        if(t_imp.equals("rep_anvenc_tot")){ 
            String hasta = request.getParameter("parFHAS"); 
            formato = "rep_anvenc_tot";
            //formato_ex = "rep_anvenc_tot_ex";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("dHASTA",hasta);
            
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            /********************PDF********************/            
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            print = jasperPrint;            
            conn.close();         
        }
        if(t_imp.equals("rep_anvenc_det")){ 
            String hasta = request.getParameter("parFHAS");
            String moneda = request.getParameter("parMON");
            formato = "rep_anvenc_det";
            formato_ex = "rep_anvenc_det_ex";                            
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("dHASTA",hasta);
            parametros.put("MONEDA",moneda);
            parametros.put("inPAR",parAdi);
            
            /********************PDF********************/  
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);                      
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;            
            conn.close();         
        }
        if(t_imp.equals("rep_ctecli")){
            formato = "rep_ctecli";
            //formato_ex = "rep_ctecli_ex";                            
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("IN_PAR",addCond);

            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            /********************PDF********************/
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/                        
            print = jasperPrint;
            conn.close();
        }
        if(t_imp.equals("rep_progra")){        
            formato = "rep_progra";
            formato_ex = "rep_progra";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();           
        }
        if(t_imp.equals("pro_ficha")){            
            formato = "pro_ficha";
            formato_ex = "pro_ficha";                            
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("CDG_FICH",nro_doc);            
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();            
        }
        if(t_imp.equals("rep_ordenm")){            
            formato = "rep_ordenm";
            formato_ex = "rep_ordenm";
                
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("CDG_ORDM",nro_doc);            
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();            
        }
        if(t_imp.equals("rep_pend")){
            String desde = request.getParameter("parFDES"); 
            String hasta = request.getParameter("parFHAS"); 
           // String c1 = request.getParameter("parCOND1");
           // String c2 = request.getParameter("parCOND2");
           // String c3 = request.getParameter("parCOND3");
          //  String opc = request.getParameter("parOPCION");
          //  String pTIT = request.getParameter("pTITULO");
            formato = "rep_pend";
           // formato_ex = "rep_pend_ex";                            
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("dDESDE",desde);
            parametros.put("dHASTA",hasta);
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar datos de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar datos de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar datos de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar datos de la empresa
            //parametros.put("cCOND1",c1);
            //parametros.put("cCOND2",c2);
            //parametros.put("cCOND3",c3);
            //parametros.put("cOPC",opc);
            parametros.put("pTITULO",pTIT);
                                    
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
         //   reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
         //   jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
         //   print = jasperPrint;
         //   conn.close();
        }
        if(t_imp.equals("rep_acep")){
            String fec_ped = request.getParameter("parFDES"); 
            String dolares = request.getParameter("parREP");
            String soles = request.getParameter("parREF");
            String ruc_cli = request.getParameter("parCOND1");
            String des_cli = request.getParameter("parCOND2");
            String importe = request.getParameter("parCOND3");
            formato = "rep_acep";
            formato_ex = "rep_acep";                            
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("cNUM_PED",nro_doc);
            parametros.put("cFEC_PED",fec_ped);
            parametros.put("cSOLES",soles);
            parametros.put("cDOLARES",dolares);
            parametros.put("cRUC_CLI",ruc_cli);
            parametros.put("cDES_CLI",des_cli);
            parametros.put("cCOND3",importe);
                                    
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();
        }
        if(t_imp.equals("rep_pago")){
            String cond = request.getParameter("addCond");
            formato = "rep_pago";
            formato_ex = "rep_pago";                            
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("cCOND1",cond);
                                    
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();
        }
        if(t_imp.equals("rep_lres")){
            String moneda = request.getParameter("parMON");
            formato = "rep_lres";
            formato_ex = "rep_lres";                            
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("pNUM_CAJ",nro_doc);
            parametros.put("pCDG_MON",moneda);
                                    
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();
        }
        if(t_imp.equals("rep_prest")){     
            String tip_emp = request.getParameter("parPROD"); 
            String cdg_emp = request.getParameter("parDESP"); 
            formato = "rep_prest";
            formato_ex = "rep_prest_ex";                            
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("cNUMPRES",nro_doc);  
            parametros.put("cCDGEMP",cdg_emp);  
            parametros.put("cTIPEMP",tip_emp);  
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();            
        }
        if(t_imp.equals("rep_cuotas")){     
            formato = "rep_cuotas";
            formato_ex = "rep_cuotas";                            
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("ID",nro_doc);   
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();            
        }
        if(t_imp.equals("rep_pla")){   
            String cANIO = request.getParameter("parCOND1"); 
            String cMES = request.getParameter("parCOND2"); 
            String cSEMANA = request.getParameter("parMON"); 
            String cTIPEMP = request.getParameter("parCOND3"); 
            String cTIPPLA = request.getParameter("parTITULO"); 
            
            formato = "rep_pla";
            formato_ex = "rep_pla_grupo";                           
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("cANIO",cANIO);
            parametros.put("cMES",cMES);
            parametros.put("cSEMANA",cSEMANA);
            parametros.put("cTIPEMP",cTIPEMP);
            parametros.put("cTIPPLA",cTIPPLA);
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            sheets.add(print);
            parametros.put(JRParameter.IS_IGNORE_PAGINATION, true);
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            sheets.add(print);                        
            conn.close();            
        }
        if(t_imp.equals("pla_lis_5ta")){   
            formato = "pla_lis_5ta";
            formato_ex = "pla_lis_5ta";                            
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("OPC",nro_doc);

            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint;
            conn.close();            
        }
        if(t_imp.equals("rep_rv_sunat")){        
            formato = "rep_rv_sunat";
            formato_ex = "rep_rv_sunat";
            
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("CDG_USU",parAdi);
            parametros.put("CDG_LOG",n_arc);
            parametros.put("RUC_USU",nro_doc);
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            /*reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);*/
            print = jasperPrint; 
            conn.close();           
        }
        
        if(t_imp.equals("rep_rvtam")){        
            formato = "rep_rvtam";
            formato_ex = "rep_rvtam";
            String address = "San Isidro Av. Salaverry 890";
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("CDG_USU",parAdi);
            parametros.put("CDG_LOG",n_arc);
            parametros.put("CDG_DIRRE",address);
            parametros.put("RUC_USU",nro_doc);
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            /*reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);*/
            print = jasperPrint; 
            conn.close();           
        }
                         
        if(t_imp.equals("rep_rvent_lineas")){        
            formato = "rep_rvent_lineas";
            formato_ex = "rep_rvent_lineas";
            String address = "San Isidro Av. Salaverry 890";
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            //parametros.put("CDG_USU",parAdi);
            parametros.put("CDG_USU",nro_doc);
            parametros.put("CDG_LOG",n_arc);
            parametros.put("CDG_DIRRE",address);
            //parametros.put("RUC_USU",nro_doc);
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            /*reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);*/
            print = jasperPrint; 
            conn.close();           
        }
        
        if(t_imp.equals("rep_rescobranza")){        
            formato = "rep_rescobranza";
            formato_ex = "rep_rescobranza_ex";
            String address = "San Isidro Av. Salaverry 890";
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("CDG_USU",parAdi);
            parametros.put("CDG_LOG",n_arc);
            parametros.put("CDG_DIRRE",address);
            parametros.put("RUC_USU",nro_doc);
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint; 
            conn.close();           
        }
        
        if(t_imp.equals("rep_ctecli_vto")){        
            formato = "rep_ctecli_vto";
            //formato_ex = "rep_rescobranza_ex";
            String address = "San Isidro Av. Salaverry 890";
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            //parametros.put("CDG_USU",parAdi);
            parametros.put("CDG_LOG",n_arc);
            parametros.put("CDG_DIRRE",address);
            parametros.put("RUC_USU",nro_doc);
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
           /* reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);*/
            print = jasperPrint; 
            conn.close();           
        }
        
        if(t_imp.equals("rep_progra")){        
            formato = "rep_progra";
            //formato_ex = "rep_rescobranza_ex";
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("CDG_USU",nro_doc);
            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
           /* reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);*/
            print = jasperPrint; 
            conn.close();           
        }
          if(t_imp.equals("rep_prov")){        
            formato = "rep_prov";
            formato_ex = "rep_prov_ex";
            
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
                
            parametros = new HashMap();
            parametros.clear();
            parametros.put("CDG_RUC",cRUC);         //agregado para mostrar datos de la empresa
            parametros.put("CDG_LOG",cNAME);        //agregado para mostrar datos de la empresa
            parametros.put("CDG_DIRE",cDIRE);       //agregado para mostrar datos de la empresa
            parametros.put("CDG_TELEF",cTELF);      //agregado para mostrar datos de la empresa

            
            /********************PDF********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            bytes = JasperExportManager.exportReportToPdf(jasperPrint);
            /********************EXCEL********************/
            reportFile = new File(application.getRealPath("/formatos/"+formato_ex+".jasper"));
            jasperPrint = JasperFillManager.fillReport(reportFile.getPath(), parametros, conn);
            print = jasperPrint; 
            conn.close();           
        }
        if(t_sal.equals("Pdf")){
            /*Indicamos que la respuesta va a ser en formato PDF*/
            response.setContentType("application/pdf");
            response.setContentLength(bytes.length);
            ServletOutputStream ouputStreamPdf = response.getOutputStream();
            ouputStreamPdf.write(bytes, 0, bytes.length);
            /*Limpiamos y cerramos flujos de salida*/
            ouputStreamPdf.flush(); 
            ouputStreamPdf.close();
        }
        if(t_sal.equals("Excel")){
            arrayOutputStream = new ByteArrayOutputStream();
            exporterXLS = new JRXlsExporter();                

            if(t_imp.equals("rep_anvenc") || t_imp.equals("rep_pla")) exporterXLS.setParameter(JRXlsExporterParameter.JASPER_PRINT_LIST, sheets);
            else exporterXLS.setParameter(JRXlsExporterParameter.JASPER_PRINT, print);
            exporterXLS.setParameter(JRXlsExporterParameter.OUTPUT_STREAM, arrayOutputStream);
            exporterXLS.setParameter(JRXlsExporterParameter.IS_ONE_PAGE_PER_SHEET, Boolean.FALSE);
            exporterXLS.setParameter(JRXlsExporterParameter.IS_DETECT_CELL_TYPE, Boolean.TRUE);
            exporterXLS.setParameter(JRXlsExporterParameter.IS_WHITE_PAGE_BACKGROUND, Boolean.FALSE);
            exporterXLS.setParameter(JRXlsExporterParameter.IS_REMOVE_EMPTY_SPACE_BETWEEN_ROWS, Boolean.TRUE);
            exporterXLS.exportReport();
            bytes = arrayOutputStream.toByteArray();
            response.setContentType("application/vnd.ms-excel");
            response.setHeader("Content-disposition", "attachment; filename="+Integer.toString(Math.abs(rd.nextInt()))+".xls");
            response.setContentLength(bytes.length);
            ouputStream = response.getOutputStream();
            ouputStream.write(bytes, 0, bytes.length);
            ouputStream.flush();
            ouputStream.close();                       
        }  
%>

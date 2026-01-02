Attribute VB_Name = "cblELPDTAQ"
Option Explicit


Public Sub ELPDTAQMON()
       IDENTIFICATION DIVISION.
        PROGRAM-ID. ELPDTAQMON.
      *---------------------------------------------------------------*-
      * S/PGM ELPDTAQMON                                              *
      *---------------------------------------------------------------*-
       ENVIRONMENT DIVISION.
        CONFIGURATION SECTION.
        SOURCE-COMPUTER. IBM-AS400.
        OBJECT-COMPUTER. IBM-AS400.
        INPUT-OUTPUT SECTION.
      *---------------------------------------------------------------*-
       FILE-CONTROL.
       DATA DIVISION.

       FILE SECTION.
      *---------------------------------------------------------------*-
       WORKING-STORAGE SECTION.
      *---------------------------------------------------------------*-


       01  QSBMJOB-TXT.
           02  FILLER            PIC X(38)
               VALUE "SBMJOB CMD(CALL PGM(ELPDTAQCL) PARM ('".
           02  QSBMJOB-DTAQ-IN  PIC X(10) VALUE SPACE.
           02  FILLER            PIC XXX   VALUE "' '".
           02  QSBMJOB-DTAQ-LIB PIC X(10) VALUE SPACE.
           02  FILLER            PIC XXX  VALUE "'))".
           02  FILLER            PIC X(13) VALUE SPACE.
      *        VALUE " JOBQ(BIASRV)".
       77  QSBMJOB-DTAQ-JOBQ  PIC X(10) VALUE SPACE.
       77  QSBMJOB-L   PIC 9(10)V9(5) COMP-3 VALUE 77.
       77  W-STATUS          PIC X(10) VALUE SPACE.
       77  APPOBJ            PIC X(10) VALUE SPACE.
       77  M-LIB             PIC X(10) VALUE SPACE.
       77  M-OUT             PIC X(10) VALUE SPACE.
       77  P5                PIC 99999 VALUE ZERO.
       01  WST.
           02  T-SRV-INDEX   PIC S999 COMP-3 VALUE ZERO.
           02  T-SRV-NB      PIC S999 COMP-3 VALUE ZERO.
           02  T-SRV-NBMAX   PIC S999 COMP-3 VALUE 10.
           02  T-SRV         OCCURS 10.
               03  T-SRV-STATUS   PIC S9 COMP-3.
               03  T-SRV-DTAQ-IN  PIC X(10).
               03  T-SRV-HEADER   PIC X(114).
               03  T-SRV-DATE     PIC 99999999.
               03  T-SRV-TIME     PIC 99999999.
           02  ELPMSG-LEN    PIC S9(5) COMP-3 VALUE 114.
           02  L-DTAQ.
               COPY DDS-ALL-FORMATS OF ELPDTAQ.
           02  L-DTAQ-LEN        PIC S9(05) COMP-3 VALUE 31744.
           02  L-DTAQ-WAIT       PIC S9(05) COMP-3 VALUE -1.

           02  W-DTAQ-NAME.
               03  FILLER         PIC XX.
               03  W-DTAQ-INDEX   PIC 999999.
               03  FILLER         PIC XX.

       01  IND-PGM.
           02  FILLER         PIC 1 VALUE B"0".
           88  FIN-OFF              VALUE B"0".
           88  FIN-ON               VALUE B"1".
           02  FILLER         PIC 1 VALUE B"0".
           88  TEST-OFF             VALUE B"0".
           88  TEST-ON              VALUE B"1".
           02  FILLER         PIC 1          VALUE B"0".
           88  LOOP-OFF             VALUE B"0".
           88  LOOP-ON              VALUE B"1".

       LINKAGE SECTION.
      *===============================================================*

       01  L-DTAQ-IN.
             02  FILLER          PIC X(10).
       01  L-DTAQ-LIB.
             02  FILLER          PIC X(10).
       01  L-DTAQ-NBX.
             02  L-DTAQ-NB       PIC 99.
       01  L-DTAQ-JOBQ.
             02  FILLER          PIC X(10).


       PROCEDURE DIVISION USING L-DTAQ-IN L-DTAQ-LIB L-DTAQ-NBX
                                L-DTAQ-JOBQ.
      *---------------------------------------------------------------*-
       PP.
      *---------------------------------------------------------------*-
           IF L-DTAQ-NB > T-SRV-NBMAX
              MOVE T-SRV-NBMAX TO T-SRV-NB
           Else
              MOVE L-DTAQ-NB   TO T-SRV-NB
           END-IF
           MOVE SPACE       TO L-DTAQ
           MOVE L-DTAQ-LIB  TO QSBMJOB-DTAQ-LIB
           MOVE L-DTAQ-JOBQ  TO QSBMJOB-DTAQ-JOBQ
           MOVE L-DTAQ-IN   TO W-DTAQ-NAME
           PERFORM ELPDTAQ-SRVSTART VARYING T-SRV-INDEX
                   FROM 1 BY 1 UNTIL        T-SRV-INDEX > T-SRV-NB

           MOVE 100000     TO W-DTAQ-INDEX.
           PERFORM ELPDTAQ-RCV  UNTIL FIN-ON.

           GO TO FIN.
      *---------------------------------------------------------------*-
       ELPDTAQ-RCV.
      *--------------------------------------------------------------- -
           CALL  "QRCVDTAQ" USING L-DTAQ-IN    L-DTAQ-LIB   L-DTAQ-LEN
                                  L-DTAQ   L-DTAQ-WAIT

           CALL  "ELPANSIRCV" USING L-DTAQ   L-DTAQ-LEN

           SET TEST-OFF TO TRUE

           EVALUATE SRVMETHOD OF L-DTAQ

               WHEN Space
                    IF T-SRV-NB = ZERO
                       PERFORM ELPDTAQ - SRV
                     Else
                       PERFORM ELPDTAQ-SRVMONITOR VARYING T-SRV-INDEX
                       FROM 1 BY 1 UNTIL   T-SRV-INDEX > T-SRV-NB
                       OR TEST-ON
                       IF TEST-OFF
                          MOVE "SRVRETRY  "   TO SRVERR OF L-DTAQ
                          MOVE L-DTAQ-LEN TO P5
                          MOVE P5         TO SRVDTAQLEN OF L-DTAQ
                          CALL  "QSNDDTAQ" USING SRVDTAQOUT OF L-DTAQ
                                                 SRVDTAQLIB OF L-DTAQ
                                                 L-DTAQ-LEN
                                                 L -DTAQ
                 END-IF
                    END-IF

               WHEN "SRVOK"
                     PERFORM ELPDTAQ-SRVOK VARYING T-SRV-INDEX
                     FROM 1 BY 1 UNTIL   T-SRV-INDEX > T-SRV-NB
                       OR TEST-ON

               WHEN "SRVSTARTOK"
                     PERFORM ELPDTAQ-SRVSTARTOK VARYING T-SRV-INDEX
                     FROM 1 BY 1 UNTIL   T-SRV-INDEX > T-SRV-NB
                       OR TEST-ON

               WHEN "SRVEND"
                     PERFORM ELPDTAQ-SRVEND       VARYING T-SRV-INDEX
                     FROM 1 BY 1 UNTIL   T-SRV-INDEX > T-SRV-NB
                       OR TEST-ON

               WHEN "PCINIT"
                     PERFORM ELPDTAQ - PCINIT

               WHEN "ELPDTAQEND"
                     PERFORM ELPDTAQ-END       VARYING T-SRV-INDEX
                     FROM 1 BY 1 UNTIL   T-SRV-INDEX > T-SRV-NB
                     MOVE ELPMSG-LEN TO P5
                     MOVE P5         TO SRVDTAQLEN OF L-DTAQ
                     CALL  "QSNDDTAQ" USING SRVDTAQOUT OF L-DTAQ
                                            SRVDTAQLIB OF L-DTAQ
                                            ELPMSG-LEN
                                            L -DTAQ
                     SET FIN-ON TO TRUE
           END-EVALUATE.

      *---------------------------------------------------------------*-
       ELPDTAQ-SRV.
      *--------------------------------------------------------------- -
           SUBTRACT ELPMSG-LEN FROM L-DTAQ-LEN
           MOVE MSGTXT OF L-DTAQ TO APPOBJ
           SET LOOP-ON TO TRUE
           PERFORM ELPDTAQ-APPOBJ UNTIL LOOP-OFF.
      *---------------------------------------------------------------*-
       ELPDTAQ-APPOBJ.
      *--------------------------------------------------------------- -


           MOVE SPACE    TO W-STATUS
           CALL APPOBJ    USING  MSGTXT OF L-DTAQ W-STATUS  L-DTAQ-LEN
           ADD ELPMSG-LEN TO L-DTAQ-LEN
           MOVE L-DTAQ-LEN TO P5
           MOVE P5         TO SRVDTAQLEN OF L-DTAQ

           CALL  "ELPANSISND" USING L-DTAQ   L-DTAQ-LEN

           CALL  "QSNDDTAQ" USING SRVDTAQOUT OF L-DTAQ
                                  SRVDTAQLIB OF L-DTAQ
                                  L-DTAQ-LEN
                                  L-DTAQ.
           IF W-STATUS NOT = "$LOOP     " SET LOOP-OFF TO TRUE END-IF.
      *----------------------------------------------------------------*
       ELPDTAQ-PCINIT.
      *----------------------------------------------------------------*

           MOVE SRVDTAQLIB OF L-DTAQ TO M-LIB
           MOVE SRVDTAQOUT OF L-DTAQ TO M-OUT
           MOVE L-DTAQ-LIB  TO SRVDTAQLIB OF L-DTAQ.
           ADD 1            TO W-DTAQ-INDEX
           MOVE W-DTAQ-NAME TO SRVDTAQOUT OF L-DTAQ.
           CALL  "ELPDTAQCZ" USING SRVDTAQOUT OF L-DTAQ
                                   SRVDTAQLIB OF L-DTAQ

           MOVE L-DTAQ-LEN TO P5
           MOVE P5         TO SRVDTAQLEN OF L-DTAQ
           CALL  "QSNDDTAQ"  USING M-OUT
                                   M -LIB
                                   L-DTAQ-LEN
                                   L-DTAQ.

      *----------------------------------------------------------------*
       ELPDTAQ-SRVOK.
      *----------------------------------------------------------------*
           IF T-SRV-DTAQ-IN (T-SRV-INDEX) = SRVDTAQIN OF L-DTAQ
              IF T-SRV-STATUS (T-SRV-INDEX) NOT = 2
                 MOVE T-SRV-HEADER (T-SRV-INDEX) TO L-DTAQ
                 IF SRVDTAQOUT OF L-DTAQ NOT = SPACE
                     MOVE "SRVERR0002"   TO SRVERR OF L-DTAQ
                     CALL  "QSNDDTAQ" USING SRVDTAQOUT OF L-DTAQ
                                            SRVDTAQLIB OF L-DTAQ
                                            ELPMSG-LEN
                                            L -DTAQ
                 END-IF
              END-IF
              MOVE ZERO TO T-SRV-STATUS (T-SRV-INDEX)
              SET TEST-ON TO TRUE
           END-IF.

      *----------------------------------------------------------------*
       ELPDTAQ-SRVSTARTOK.
      *----------------------------------------------------------------*
           IF T-SRV-DTAQ-IN (T-SRV-INDEX) = SRVDTAQIN OF L-DTAQ
              MOVE ZERO TO T-SRV-STATUS (T-SRV-INDEX)
              SET TEST-ON TO TRUE
           END-IF.

      *----------------------------------------------------------------*
       ELPDTAQ-SRVEND.
      *----------------------------------------------------------------*
           IF T-SRV-DTAQ-IN (T-SRV-INDEX) = SRVDTAQIN OF L-DTAQ
              MOVE 9 TO T-SRV-STATUS (T-SRV-INDEX)
              SET TEST-ON TO TRUE
           END-IF.

      *----------------------------------------------------------------*
       ELPDTAQ-SRVMONITOR.
      *----------------------------------------------------------------*
           IF T-SRV-STATUS (T-SRV-INDEX) = ZERO
              MOVE 2 TO T-SRV-STATUS (T-SRV-INDEX)
              ACCEPT T-SRV-DATE (T-SRV-INDEX) FROM DATE
              ACCEPT T-SRV-TIME (T-SRV-INDEX) FROM TIME
              MOVE L-DTAQ TO T-SRV-HEADER (T-SRV-INDEX)
              CALL "QSNDDTAQ" USING T-SRV-DTAQ-IN (T-SRV-INDEX)
                                    L -DTAQ - LIB
                                    L-DTAQ-LEN
                                    L -DTAQ
              SET TEST-ON TO TRUE
           END-IF.
      *----------------------------------------------------------------*
       ELPDTAQ-SRVSTART.
      *----------------------------------------------------------------*
           ADD 1            TO W-DTAQ-INDEX
           MOVE W-DTAQ-NAME TO QSBMJOB-DTAQ-IN
                               T-SRV-DTAQ-IN (T-SRV-INDEX)
           MOVE 1           TO T-SRV-STATUS (T-SRV-INDEX)
           ACCEPT T-SRV-DATE (T-SRV-INDEX) FROM DATE
           ACCEPT T-SRV-TIME (T-SRV-INDEX) FROM TIME
           MOVE L-DTAQ      TO T-SRV-HEADER (T-SRV-INDEX)
           CALL  "ELPDTAQCB" USING QSBMJOB-DTAQ-IN
                                   QSBMJOB -DTAQ - LIB
                                   QSBMJOB-DTAQ-JOBQ.

      *----------------------------------------------------------------*
       ELPDTAQ-END.
      *----------------------------------------------------------------*
           IF T-SRV-STATUS (T-SRV-INDEX) NOT = 9
              MOVE 9 TO T-SRV-STATUS (T-SRV-INDEX)
              CALL "QSNDDTAQ" USING T-SRV-DTAQ-IN (T-SRV-INDEX)
                                    L -DTAQ - LIB
                                    L-DTAQ-LEN
                                    L -DTAQ
           END-IF.

      *===============================================================*
       FIN.
           EXIT PROGRAM.
      *===============================================================*

End Sub

Public Sub ELPDTAQUSR()
      IDENTIFICATION DIVISION.
        PROGRAM-ID. ELPDTAQUSR.
      *---------------------------------------------------------------*-
      * S/PGM ELPDTAQUSR                                              *
      *---------------------------------------------------------------*-
       ENVIRONMENT DIVISION.
        CONFIGURATION SECTION.
        SOURCE-COMPUTER. IBM-AS400.
        OBJECT-COMPUTER. IBM-AS400.
        INPUT-OUTPUT SECTION.
      *---------------------------------------------------------------*-
       FILE-CONTROL.
       DATA DIVISION.

       FILE SECTION.
      *---------------------------------------------------------------*-
       WORKING-STORAGE SECTION.
      *---------------------------------------------------------------*-

      *!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
       77  NB-MAX            PIC S999 COMP-3 VALUE 10.
       77  SRV-EXBIDT-LEN      PIC S999 COMP-3 VALUE 49.
      *!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
       77  W-STATUS          PIC X(10) VALUE SPACE.
       77  X-STATUS          PIC X(10) VALUE SPACE.
       77  SP-PGM            PIC X(10) VALUE SPACE.
       01  WST.
           02  W-EXBIDT.
               COPY DDS-ALL-FORMATS OF EXBIDTP0.
       01  IND-PGM.
           02  FILLER    PIC 1      VALUE B"0".
           88  TEST-OFF             VALUE B"0".
           88  TEST-ON              VALUE B"1".

       LINKAGE SECTION.
      *===============================================================*
       01  L-MSGTXT.
      *
           02  W-APP-OBJET          PIC X(12).
           02  W-APP-METHOD         PIC X(12).
           02  W-APP-ERR            PIC X(10).
           02  W-PC-ID              PIC X(10).
           02  W-RTN.
                04  W-SOC-ID             PIC 9(03).
                04  W-SOC-AGENCE         PIC 9(03).
                04  W-SOC-BDFG2          PIC 9(05).
                04  W-SOC-BDFG3          PIC 9(05).
                04  W-SOC-NAME           PIC X(40).
                04  W-USR-NB             PIC 9(02).
                04  W-USR                OCCURS 10.
                   05  W-USR-ID             PIC X(10).
                   05  W-USR-SERV           PIC 999.
                   05  W-USR-COGES          PIC 99.
                   05  W-USR-TYPER          PIC X(04).
                   05  W-USR-PRPER          PIC X(15).
                   05  W-USR-NOPER          PIC X(15).
       01  L-STATUS.
           02  L-STATUS-TXT      PIC X(8).
           02  L-STATUS-CODE-N   PIC 99.
       01  L-DTAQ-LEN-X.
           02  L-DTAQ-LEN        PIC S9(5) COMP-3.

       PROCEDURE DIVISION USING L-MSGTXT    L-STATUS L-DTAQ-LEN-X.
      *===============================================================*
      *---------------------------------------------------------------*-
       PP.
      *---------------------------------------------------------------*-
           MOVE SPACE    TO W-RTN  L-STATUS
           MOVE ZERO     TO IND-PGM
                            W -USR - Nb
           MOVE "B INTERCONT ARABE PARIS" TO W-SOC-NAME
           MOVE 12179          TO W-SOC-BDFG2
           MOVE 1              TO W-SOC-ID  W-SOC-AGENCE W-SOC-BDFG3
           MOVE 102 TO L-DTAQ-LEN

           EVALUATE W - App - Method
             WHEN "PCID        "
               PERFORM EXBIDT - pcId
             WHEN "USRID       "
               PERFORM EXBIDT - usrId

             WHEN OTHER
                MOVE "METHOD"      TO W-APP-ERR
           END-EVALUATE.

           GoTo FIN
      *===============================================================*
           EXIT PROGRAM.
      *===============================================================*
      *----------------------------------------------------------------*
       EXBIDT-USRID.
      *----------------------------------------------------------------*

           MOVE "EXBIDTP0R" TO SP-PGM
           MOVE SPACE  TO W-STATUS
           MOVE W-PC-ID OF L-MSGTXT  TO NOMOP    OF W-EXBIDT
           CALL SP-PGM    USING W-EXBIDT W-STATUS.
           PERFORM SRV-USR.
      *----------------------------------------------------------------*
       EXBIDT-PCID.
      *----------------------------------------------------------------*

           MOVE "EXBIDTL3R" TO SP-PGM
           MOVE "SREQ" TO W-STATUS
           MOVE W-PC-ID OF L-MSGTXT  TO ECRAN    OF W-EXBIDT
           MOVE  SPACE           TO NOMOP    OF W-EXBIDT
           CALL SP-PGM    USING W-EXBIDT W-STATUS.
           PERFORM EXBIDT-SNAP-RN UNTIL W-STATUS NOT = SPACE.
      *----------------------------------------------------------------*
       EXBIDT-SNAP-RN.
      *----------------------------------------------------------------*
           MOVE "RNEQ"  TO W-STATUS
           CALL SP-PGM    USING W-EXBIDT   W-STATUS.
           PERFORM SRV-USR.
      *----------------------------------------------------------------*
       SRV-USR.
      *----------------------------------------------------------------*
           IF W-STATUS = SPACE
              ADD SRV-EXBIDT-LEN TO L-DTAQ-LEN
              ADD 1 TO W-USR-NB
              MOVE NOMOP    OF W-EXBIDT TO W-USR-ID    (W-USR-NB)
              MOVE SERV     OF W-EXBIDT TO W-USR-SERV  (W-USR-NB)
              MOVE COGES    OF W-EXBIDT TO W-USR-COGES (W-USR-NB)
              MOVE TYPER    OF W-EXBIDT TO W-USR-TYPER (W-USR-NB)
              MOVE PRPER    OF W-EXBIDT TO W-USR-PRPER (W-USR-NB)
              MOVE NOPER    OF W-EXBIDT TO W-USR-NOPER (W-USR-NB)


              IF W-USR-NB = NB-MAX     MOVE "SUITE" TO W-STATUS END-IF
           END-IF.
      *===============================================================*
       FIN.
           EXIT PROGRAM.
      *===============================================================*

End Sub

Public Sub ELPDTAQSND()
    IDENTIFICATION DIVISION.
        PROGRAM-ID. ELPDTAQSND.
      *---------------------------------------------------------------*-
      * S/PGM ELPDTAQSND                                              *
      *
      *---------------------------------------------------------------*-
       ENVIRONMENT DIVISION.
        CONFIGURATION SECTION.
        SOURCE-COMPUTER. IBM-AS400.
        OBJECT-COMPUTER. IBM-AS400.
        INPUT-OUTPUT SECTION.
      *---------------------------------------------------------------*-
       FILE-CONTROL.
       DATA DIVISION.

       FILE SECTION.
      *---------------------------------------------------------------*-
       WORKING-STORAGE SECTION.
      *---------------------------------------------------------------*-
       77  W-STATUS          PIC X(10) VALUE SPACE.
       77  APPOBJ            PIC X(10) VALUE SPACE.

       01  WST.
           02  ELPMSG-LEN         PIC S9(5) COMP-3 VALUE 104.
           02  L-DTAQ.
               COPY DDS-ALL-FORMATS OF ELPDTAQ.
           02  L-DTAQ-LEN        PIC S9(05) COMP-3 VALUE 31744.
           02  L-DTAQ-WAIT       PIC S9(05) COMP-3 VALUE -1.


       01  IND-PGM.
           02  FILLER         PIC 1          VALUE B"0".
           88  FIN-OFF             VALUE B"0".
           88  FIN-ON              VALUE B"1".

       LINKAGE SECTION.
      *===============================================================*

           01  L-DTAQ-LIB.
             02  FILLER          PIC X(10).
           01  L-DTAQ-IN.
             02  FILLER          PIC X(10).
           01  L-DTAQ-FCT.
             02  FILLER          PIC X(10).


       PROCEDURE DIVISION USING   L-DTAQ-IN  L-DTAQ-LIB L-DTAQ-FCT.
      *---------------------------------------------------------------*-
       PP.
      *---------------------------------------------------------------*-

           MOVE SPACE        TO L-DTAQ
           MOVE L-DTAQ-IN    TO SRVDTAQOUT OF L-DTAQ
           MOVE L-DTAQ-LIB   TO SRVDTAQLIB OF L-DTAQ

           IF L-DTAQ-FCT = "ELPDTAQEND"
              MOVE "ELPDTAQEND" TO SRVMETHOD OF L-DTAQ

              CALL  "QSNDDTAQ" USING SRVDTAQOUT OF L-DTAQ
                                     SRVDTAQLIB OF L-DTAQ
                                     L-DTAQ-LEN
                                     L -DTAQ
           END-IF.

      *===============================================================*
       FIN.
           EXIT PROGRAM.
      *===============================================================*

End Sub

Public Sub ELPDTAQ()
  IDENTIFICATION DIVISION.
        PROGRAM-ID. ELPDTAQ.
      *---------------------------------------------------------------*-
      * S/PGM ELPDTAQ                                                 *
      *
      *---------------------------------------------------------------*-
       ENVIRONMENT DIVISION.
        CONFIGURATION SECTION.
        SOURCE-COMPUTER. IBM-AS400.
        OBJECT-COMPUTER. IBM-AS400.
        INPUT-OUTPUT SECTION.
      *---------------------------------------------------------------*-
       FILE-CONTROL.
       DATA DIVISION.

       FILE SECTION.
      *---------------------------------------------------------------*-
       WORKING-STORAGE SECTION.
      *---------------------------------------------------------------*-
       77  W-STATUS          PIC X(10) VALUE SPACE.
       77  APPOBJ            PIC X(10) VALUE SPACE.
       77  P5                PIC 99999 VALUE ZERO.

       01  WST.
           02  ELPMSG-LEN         PIC S9(5) COMP-3 VALUE 114.
           02  L-DTAQ.
               COPY DDS-ALL-FORMATS OF ELPDTAQ.
           02  L-DTAQ-LEN        PIC S9(05) COMP-3 VALUE 31744.
           02  L-DTAQ-WAIT       PIC S9(05) COMP-3 VALUE -1.

           02  MON-DTAQ.
               COPY DDS-ALL-FORMATS OF ELPDTAQ.
           02  W-DTAQ-IN.
               03  FILLER          PIC XX.
               03  W-DTAQ-INDEX    PIC 999999.
               03  FILLER          PIC XX.

       01  IND-PGM.
           02  FILLER         PIC 1          VALUE B"0".
           88  FIN-OFF             VALUE B"0".
           88  FIN-ON              VALUE B"1".
           02  FILLER         PIC 1          VALUE B"0".
           88  LOOP-OFF             VALUE B"0".
           88  LOOP-ON              VALUE B"1".

       LINKAGE SECTION.
      *===============================================================*

       01  L-DTAQ-LIB.
             02  FILLER          PIC X(10).
       01  L-DTAQ-IN.
             02  FILLER          PIC X(10).


       PROCEDURE DIVISION USING   L-DTAQ-IN  L-DTAQ-LIB.
      *---------------------------------------------------------------*-
       PP.
      *---------------------------------------------------------------*-
           PERFORM ELPDTAQ - STARTOK
           PERFORM ELPDTAQ-SRV  UNTIL FIN-ON.

           GO TO FIN.
      *---------------------------------------------------------------*-
       ELPDTAQ-SRV.
      *--------------------------------------------------------------- -
           CALL  "QRCVDTAQ" USING L-DTAQ-IN    L-DTAQ-LIB   L-DTAQ-LEN
                                  L-DTAQ   L-DTAQ-WAIT

           IF SRVMETHOD OF L-DTAQ = "ELPDTAQEND"
              PERFORM ELPDTAQ - SrvMethod
           Else

              SUBTRACT ELPMSG-LEN  FROM L-DTAQ-LEN
              MOVE MSGTXT OF L-DTAQ TO APPOBJ
              SET LOOP-ON TO TRUE
              PERFORM ELPDTAQ-APPOBJ UNTIL LOOP-OFF
           END-IF.

           CALL  "QSNDDTAQ"  USING SRVDTAQOUT OF MON-DTAQ
                                   SRVDTAQLIB OF MON-DTAQ
                                   ELPMSG-LEN
                                   MON-DTAQ.
      *---------------------------------------------------------------*-
       ELPDTAQ-APPOBJ.
      *--------------------------------------------------------------- -

           MOVE SPACE    TO W-STATUS
           CALL APPOBJ  USING  MSGTXT OF L-DTAQ W-STATUS  L-DTAQ-LEN
           ADD ELPMSG-LEN TO L-DTAQ-LEN
           MOVE L-DTAQ-LEN TO P5
           MOVE P5         TO SRVDTAQLEN OF L-DTAQ

           CALL  "ELPANSISND" USING L-DTAQ   L-DTAQ-LEN

           CALL  "QSNDDTAQ" USING SRVDTAQOUT OF L-DTAQ
                                  SRVDTAQLIB OF L-DTAQ
                                  L-DTAQ-LEN
                                  L -DTAQ
           IF W-STATUS NOT = "$LOOP     " SET LOOP-OFF TO TRUE END-IF.

      *---------------------------------------------------------------*-
       ELPDTAQ-STARTOK.
      *--------------------------------------------------------------- -

           MOVE SPACE        TO MON-DTAQ
           MOVE L-DTAQ-IN    TO W-DTAQ-IN
           MOVE 1            TO W-DTAQ-INDEX
           MOVE W-DTAQ-IN    TO SRVDTAQOUT OF MON-DTAQ
           MOVE L-DTAQ-LIB   TO SRVDTAQLIB OF MON-DTAQ
           MOVE L-DTAQ-IN    TO SRVDTAQIN  OF MON-DTAQ
           MOVE "SRVSTARTOK" TO SRVMETHOD  OF MON-DTAQ

              MOVE ELPMSG-LEN TO P5
              MOVE P5         TO SRVDTAQLEN OF L-DTAQ
           CALL  "QSNDDTAQ"  USING SRVDTAQOUT OF MON-DTAQ
                                   SRVDTAQLIB OF MON-DTAQ
                                   ELPMSG-LEN
                                   MON-DTAQ.
           MOVE "SRVOK"      TO SRVMETHOD  OF MON-DTAQ.
      *---------------------------------------------------------------*-
       ELPDTAQ-SRVMETHOD.
      *--------------------------------------------------------------- -
           SET FIN-ON TO TRUE
           MOVE "SRVEND"     TO SRVMETHOD  OF MON-DTAQ.
      *===============================================================*
       FIN.
           MOVE SPACE        TO L-DTAQ-IN
           EXIT PROGRAM.
      *===============================================================*

End Sub

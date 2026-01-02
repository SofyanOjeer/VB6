Attribute VB_Name = "clpELPDTAQ"
Option Explicit


Public Sub STRDTAQTST()
/*********************************************************************/
/*                                                                   */
/*  STRDTAQTST :                                                     */
/*                                                                   */
/*********************************************************************/

             PGM

             DCL        VAR(&DATPRO) TYPE(*CHAR) LEN(6)
             DCL        VAR(&DATJOB) TYPE(*CHAR) LEN(6)

             RTVJOBA    DATE(&DATJOB)

             RTVDTAARA  DTAARA(BIAFIL/DATAPRO) RTNVAR(&DATPRO)


             IF         COND(&DATPRO *NE &DATJOB) THEN(DO) /* +
                          Démarrage planifié pour le lendemain +
                          matin) */

             SBMJOB     CMD(CALL PGM(ELPDTAQTST)) JOB(DTAQTST) +
                          JOBQ(QPGMR) SCDDATE(&DATPRO) SCDTIME(070200)

             MONMSG     MSGID(CPF1634) EXEC(DO) /* Date et heure +
                          déjà dépassée */

             SBMJOB     CMD(CALL PGM(ELPDTAQTST)) JOB(DTAQTST) +
                          JOBQ (QPGMR)

             ENDDO      /* Fin CPF1634 */

             MONMSG     MSGID(CPF1338) EXEC(DO) /* Date et heure +
                          déjà dépassée */

             SBMJOB     CMD(CALL PGM(ELPDTAQTST)) JOB(DTAQTST) +
                          JOBQ (QPGMR)

             ENDDO      /* Fin CPF1338 */

             ENDDO



             IF         COND(&DATPRO = &DATJOB) THEN(DO) /* Démarrge +
                          immédiat */

             SBMJOB     CMD(CALL PGM(ELPDTAQTST)) JOB(DTAQTST) +
                          JOBQ (QPGMR)
             ENDDO



             ENDPGM

End Sub

Public Sub STRDTAQCL()
/*********************************************************************/
/*                                                                   */
/*  STRDTAQCL:    DEMARRAGE PLANIFIE DU SOUS-SYSTEME BIASRV          */
/*                                                                   */
/*********************************************************************/

             PGM

             DCL        VAR(&DATPRO) TYPE(*CHAR) LEN(6)
             DCL        VAR(&DATJOB) TYPE(*CHAR) LEN(6)

             RTVJOBA    DATE(&DATJOB)

             RTVDTAARA  DTAARA(BIAFIL/DATAPRO) RTNVAR(&DATPRO)


             IF         COND(&DATPRO *NE &DATJOB) THEN(DO) /* +
                          Démarrage planifié pour le lendemain +
                          matin) */

             SBMJOB     CMD(CALL PGM(ELPDTAQSRV)) JOB(DTAQSRV) +
                          JOBQ(BIASRV) SCDDATE(&DATPRO) SCDTIME(070000)

             MONMSG     MSGID(CPF1634) EXEC(DO) /* Date et heure +
                          déjà dépassée */

             SBMJOB     CMD(CALL PGM(ELPDTAQSRV)) JOB(DTAQSRV) +
                          JOBQ (BIASRV)

             ENDDO      /* Fin CPF1634 */

             MONMSG     MSGID(CPF1338) EXEC(DO) /* Date et heure +
                          déjà dépassée */

             SBMJOB     CMD(CALL PGM(ELPDTAQSRV)) JOB(DTAQSRV) +
                          JOBQ (BIASRV)

             ENDDO      /* Fin CPF1338 */

             ENDDO



             IF         COND(&DATPRO = &DATJOB) THEN(DO) /* Démarrge +
                          immédiat */

             SBMJOB     CMD(CALL PGM(ELPDTAQSRV)) JOB(DTAQSRV) +
                          JOBQ (BIASRV)
             ENDDO


             ENDPGM

End Sub

Public Sub ELPDTAQxx()
/*********************************************************************/
/*  ELPDTAQJPL: GESTION DES ECHANGES AS400 / PC                      */
/*              PAR 'DTAQ'                                           */
/*********************************************************************/
             PGM

             DCL        VAR(&DTAQNB) TYPE(*CHAR) LEN(02) VALUE('00')

             DCL        VAR(&DTAQLIB) TYPE(*CHAR) LEN(10) +
                          VALUE('BIADTAQ   ')

             DCL        VAR(&DTAQIN) TYPE(*CHAR)  LEN(10) +
                          VALUE('XX000001  ')

             DCL        VAR(&DTAQOUT) TYPE(*CHAR)  LEN(10) +
                          VALUE('XX000000  ')


             DCL        VAR(&JOBQ) TYPE(*CHAR) LEN(10) +
                          VALUE('QPGMR     ')


             CALL       PGM(ELPDTAQCZ) PARM(&DTAQIN  &DTAQLIB)
             CALL       PGM(ELPDTAQCZ) PARM(&DTAQOUT &DTAQLIB)

             CALL       PGM(ELPDTAQMON) PARM(&DTAQIN &DTAQLIB +
                          &DTAQNB &JOBQ)


             ENDPGM

End Sub

Public Sub ELPDTAQTST()
/*********************************************************************/
/*  ELPDTAQJPL: GESTION DES ECHANGES AS400 / PC                      */
/*              PAR 'DTAQ'                                           */
/*********************************************************************/
             PGM

             DCL        VAR(&DTAQNB) TYPE(*CHAR) LEN(02) VALUE('02')

             DCL        VAR(&DTAQLIB) TYPE(*CHAR) LEN(10) +
                          VALUE('BIADTAQ   ')

             DCL        VAR(&JOBQ) TYPE(*CHAR) LEN(10) +
                          VALUE('QPGMR     ')

             DCL        VAR(&DTAQIN) TYPE(*CHAR)  LEN(10) +
                          VALUE('XX000001  ')

             DCL        VAR(&DTAQOUT) TYPE(*CHAR)  LEN(10) +
                          VALUE('XX000000  ')
             RMVLIBLE LIB(BIAFIL)
             MONMSG MSGID(CPF2104)
             RMVLIBLE LIB(BIATST)
             MONMSG MSGID(CPF2104)
             ADDLIBLE   LIB(BIATST) POSITION(*FIRST)
             MONMSG MSGID(CPF2103)
             ADDLIBLE   LIB(&DTAQLIB) POSITION(*LAST)
             MONMSG MSGID(CPF2103)

             CALL       PGM(ELPDTAQCZ) PARM(&DTAQIN  &DTAQLIB)
             CALL       PGM(ELPDTAQCZ) PARM(&DTAQOUT &DTAQLIB)

             CALL       PGM(ELPDTAQMON) PARM(&DTAQIN &DTAQLIB +
                          &DTAQNB &JOBQ)


             ENDPGM

End Sub

Public Sub ELPDTAQSRV()
/*********************************************************************/
/*  ELPDTAQJPL: GESTION DES ECHANGES AS400 / PC                      */
/*              PAR 'DTAQ'                                           */
/*********************************************************************/
             PGM

             DCL        VAR(&DTAQNB) TYPE(*CHAR) LEN(02) VALUE('03')

             DCL        VAR(&DTAQLIB) TYPE(*CHAR) LEN(10) +
                          VALUE('BIADTAQ   ')

             DCL        VAR(&JOBQ) TYPE(*CHAR) LEN(10) +
                          VALUE('BIASRV    ')

             DCL        VAR(&DTAQIN) TYPE(*CHAR)  LEN(10) +
                          VALUE('PC000001  ')

             DCL        VAR(&DTAQOUT) TYPE(*CHAR)  LEN(10) +
                          VALUE('PC000000  ')

             MONMSG MSGID(CPF2103)
             ADDLIBLE   LIB(&DTAQLIB) POSITION(*LAST)

             CALL       PGM(ELPDTAQCZ) PARM(&DTAQIN  &DTAQLIB)
             CALL       PGM(ELPDTAQCZ) PARM(&DTAQOUT &DTAQLIB)

             CALL       PGM(ELPDTAQMON) PARM(&DTAQIN &DTAQLIB +
                          &DTAQNB &JOBQ)


             ENDPGM

End Sub

Public Sub ELPDTAQEND()
/*********************************************************************/
/*  ELPDTAQEND: GESTION DES ECHANGES AS400 / PC                      */
/*              PAR 'DTAQ'  : ARRET DES PROGRAMMES 'SERVEUR'         */
/*********************************************************************/
             PGM

             DCL        VAR(&METHOD) TYPE(*CHAR) LEN(10) +
                          VALUE('ELPDTAQEND')

             DCL        VAR(&DTAQLIB) TYPE(*CHAR) LEN(10) +
                          VALUE('BIADTAQ   ')

             DCL        VAR(&DTAQIN) TYPE(*CHAR)  LEN(10) +
                          VALUE('PC000001  ')


             CALL       PGM(ELPDTAQSND) PARM(&DTAQIN &DTAQLIB &METHOD)


             ENDPGM

End Sub

Public Sub ELPDTAQCZ()
/*********************************************************************/
/*  ELPDTAQCZ : GESTION DES ECHANGES AS400 / PC                      */
/*              PAR 'DTAQ'                                           */
/*********************************************************************/
             PGM        PARM(&DTAQIN  &DTAQLIB)

             DCL        VAR(&DTAQLIB) TYPE(*CHAR)  LEN(10)

             DCL        VAR(&DTAQIN) TYPE(*CHAR)  LEN(10)

             MONMSG MSGID(CPF0000)

             DLTDTAQ    DTAQ(&DTAQLIB/&DTAQIN)
             CRTDTAQ    DTAQ(&DTAQLIB/&DTAQIN) MAXLEN(31744)


             ENDPGM

End Sub

Public Sub ELPDTAQCL()
/*********************************************************************/
/*  ELPDTAQCL : GESTION DES ECHANGES AS400 / PC                      */
/*              PAR 'DTAQ'                                           */
/*********************************************************************/
             PGM        PARM(&DTAQIN  &DTAQLIB)


             DCL        VAR(&JOBQ) TYPE(*CHAR)  LEN(10) +
                          VALUE('BIASRV    ')
             DCL        VAR(&DTAQLIB) TYPE(*CHAR)  LEN(10)

             DCL        VAR(&DTAQIN) TYPE(*CHAR)  LEN(10) +

             MONMSG MSGID(CPF0000)


AGAIN:
             CALL PGM(ELPDTAQ) PARM(&DTAQIN &DTAQLIB)

             IF         COND(&DTAQIN *NE '          ') THEN(DO)
             GOTO       CMDLBL(AGAIN)
             ENDDO

             ENDPGM

End Sub

Public Sub ELPDTAQCB()
/*********************************************************************/
/*  ELPDTAQCB : GESTION DES ECHANGES AS400 / PC                      */
/*              PAR 'DTAQ'                                           */
/*********************************************************************/
             PGM        PARM(&DTAQIN  &DTAQLIB &JOBQ)

             DCL        VAR(&DTAQLIB) TYPE(*CHAR)  LEN(10)

             DCL        VAR(&DTAQIN) TYPE(*CHAR)  LEN(10)
             DCL        VAR(&JOBQ) TYPE(*CHAR)  LEN(10)

             MONMSG MSGID(CPF0000)

             CALL       PGM(ELPDTAQCZ) PARM(&DTAQIN  &DTAQLIB)

             SBMJOB     CMD(CALL PGM(ELPDTAQCL) PARM(&DTAQIN +
                          &DTAQLIB)) JOB(&DTAQIN) JOBQ(&JOBQ)
/*                        &DTAQLIB)) JOB(&DTAQIN) JOBQ(BIASRV)  */

FINPRO:      ENDPGM

End Sub

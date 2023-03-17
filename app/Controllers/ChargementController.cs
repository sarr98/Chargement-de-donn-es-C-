using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using app.Models;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Bibliography;


namespace app.Controllers
{
    public class ChargementController : Controller
    {
       
        // GET: Chargement
        public ActionResult Index()
        {
            try {
                string mois = ""; 
                string annee = ""; 
                string idPole = "";
                string codeFamille = "";
                //if (variable != null)
                //{
                //    mois = variable.mois;
                //    annee = variable.annee;
                //    idPole = variable.idPole;
                //}
                DAL dal = new DAL();
                List<SelectListItem> listMois = dal.FcGetListMois();
                List<SelectListItem> listAnnee = dal.FcGetListAnnee();
                List<SelectListItem> listNomPole = dal.FcGetListePoles();
                List<SelectListItem> listFamillePole = dal.FcGetFamillePoles();
                ViewBag.ListMois = listMois;
                ViewBag.ListAnnee = listAnnee;
                ViewBag.ListNomPoles = listNomPole;
                ViewBag.listFamillePole= listFamillePole;
                ViewBag.ErrorMessage = null;
                //if (string.IsNullOrEmpty(mois) || string.IsNullOrEmpty(annee) || string.IsNullOrEmpty(idPole))
                //{
                //    //ViewBag.ErrorMessage = "Veuillez saisir les informations nécessaires.";
                //    throw new Exception("Veuillez saisir les informations nécessaires.");
                //    //return View();
                //}
                //int intMois, intAnnee, intNomPole;
                //if (!int.TryParse(mois, out intMois) || !int.TryParse(annee, out intAnnee) || !int.TryParse(idPole, out intNomPole))
                //{
                //    //ViewBag.ErrorMessage = "Les informations saisies sont incorrectes.";
                //    throw new Exception("Les informations saisies sont incorrectes.");
                //    //return View();
                //}
                //if (variable == null)
                //    variable = new Variable();

                return View(new Variable());
            }

            catch (Exception ex) {
                ViewBag.ErrorMessage = ex.Message;
                return View(new Variable());
            }
           
        }

        [HttpPost]
        public ActionResult Index(Variable variable)
        {
            try
            {
                string mois = "";
                string annee = "";
                string idPole = "";
                string codeFamille = "";
                string nomPole = "";
                string nomFichier = "";

                if (variable == null)
                    return View(variable);


                HttpPostedFileBase file = Request.Files["file1"];
                mois = variable.mois;
                annee = variable.annee;
                idPole = variable.idPole;
                codeFamille = variable.CodeFamille;
                
                DAL dal = new DAL();
                List<SelectListItem> listMois = dal.FcGetListMois();
                List<SelectListItem> listAnnee = dal.FcGetListAnnee();
                List<SelectListItem> listNomPole = dal.FcGetListePoles();
                List<SelectListItem> listFamillePole = dal.FcGetFamillePoles();
                ViewBag.ListMois = listMois;
                ViewBag.ListAnnee = listAnnee;
                ViewBag.ListNomPoles = listNomPole;
                ViewBag.listFamillePole = listFamillePole;
                ViewBag.ErrorMessage = null;


                bool chargementReussi = false;

                try
                {
                    dal.ChargerDonnees(file, codeFamille, idPole, Int32.Parse(mois), Int32.Parse(annee));
                    chargementReussi = true;
                }
                catch (Exception ex)
                {
                    ViewBag.ErrorMessage = ex.Message;
                }

                if (chargementReussi)
                {
                    ViewBag.SuccessMessage = "Les données ont été chargées avec succès.";
                }

                return View(variable);

            }

            catch (Exception ex)
            {
                ViewBag.ErrorMessage = ex.Message;
                return View(variable);
            }

        }

        





    }



}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Web.Mvc;
using System.Configuration;
using OfficeOpenXml;
using System.IO;
namespace app.Models
{
    public class DAL
    {
        public DataTable GetData(string query)
        {
            string constring = ConfigurationManager.ConnectionStrings["AdminConn"].ToString();


            try
            {
                SqlConnection sqlCon = new SqlConnection(constring);
                if (sqlCon.State == ConnectionState.Closed) { sqlCon.Open(); }
                DataSet tempDS = new DataSet();
                SqlDataAdapter tempDA = new SqlDataAdapter(string.Format(query), sqlCon);
                tempDA.Fill(tempDS, "ResultSet");
                DataTable resultSet = tempDS.Tables["ResultSet"];

                sqlCon.Close();

                return resultSet;

            }
            catch (Exception ex)
            {
                throw new Exception("GetData: La requete suivante a ramené NULL: \n query=" + query + "\n Erreur =>" + ex.Message);              
            }

        }


        //Fonction pour recevoir la liste des mois
        public List<SelectListItem> FcGetListMois()
        {
            List<SelectListItem> listMois = new List<SelectListItem>()
                {
                    new SelectListItem {Value = "1", Text = "Janvier"},
                    new SelectListItem {Value = "2", Text = "Février"},
                    new SelectListItem {Value = "3", Text = "Mars"},
                    new SelectListItem {Value = "4", Text = "Avril"},
                    new SelectListItem {Value = "5", Text = "Mai"},
                    new SelectListItem {Value = "6", Text = "Juin"},
                    new SelectListItem {Value = "7", Text = "Juillet"},
                    new SelectListItem {Value = "8", Text = "Août"},
                    new SelectListItem {Value = "9", Text = "Septembre"},
                    new SelectListItem {Value = "10", Text = "Octobre"},
                    new SelectListItem {Value = "11", Text = "Novembre"},
                    new SelectListItem {Value = "12", Text = "Décembre"},
                };
            return listMois;
        }
        public List<SelectListItem> FcGetFamillePoles()
        {
            List<SelectListItem> listFamillePole = new List<SelectListItem>()
                {
                    new SelectListItem {Value = "BQ", Text = "Banques"},
                    new SelectListItem {Value = "PP", Text = "Pôles Publiques"},
                    new SelectListItem {Value = "CO", Text = "Cotecna"},
                    new SelectListItem {Value = "AC", Text = " Assurances / Courtiers"},
                    new SelectListItem {Value = "OB", Text = "ORBUS2000"},
                   
                };
           
            return listFamillePole;
        }

        //Fonction pour recevoir la liste des année
        public List<SelectListItem> FcGetListAnnee()
        {
            int anneeActuelle = DateTime.Now.Year;
            List<SelectListItem> listAnnee = new List<SelectListItem>();
            SelectListItem a = new SelectListItem();
            a = new SelectListItem() { Value = "", Text = "Choisir l'année" };
            listAnnee.Add(a);

            for (int i = 2000; i< anneeActuelle; i++)
            {
                
                listAnnee.Add(new SelectListItem { Value = i.ToString(), Text = i.ToString() });
            }
            return listAnnee;
        }

        //Fonction pour recevoir la liste des poles
        public List<SelectListItem> FcGetListePoles()
        {
            string query = "SELECT  NOMPOLE FROM POLEMONITORING  ORDER BY NOMPOLE ASC";
            DataTable dt = GetData(query);
            List<SelectListItem> maListe = new List<SelectListItem>();
            SelectListItem a = new SelectListItem();
            a = new SelectListItem() { Value = "", Text = "Choisir le nom du pole" };
            maListe.Add(a);
            foreach (DataRow row in dt.Rows)
            {
                maListe.Add(new SelectListItem() { Value = row["NOMPOLE"].ToString(), Text = row["NOMPOLE"].ToString() });
            } 

            return maListe;
        }
        public void ChargerDonnees(HttpPostedFileBase file , string CodeFamille, string NomPole, int Periode, int Annee)
        {

            SqlConnection connection = null;
            string constring = ConfigurationManager.ConnectionStrings["AdminConn"].ToString();
            try
            {
                
                
                // Connexion à la base de données
                connection = new SqlConnection(constring);
                connection.Open();
                {
                    // Vérification de l'extension du fichier
                    string fileExtension = Path.GetExtension(file.FileName);
                    if (fileExtension != ".xls" && fileExtension != ".xlsx")
                    {
                        throw new ArgumentException("Le fichier doit être un fichier Excel (.xls ou .xlsx).");
                    }


                    using (ExcelPackage package = new ExcelPackage(file.InputStream))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                        // Récupération de la feuille de calcul
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                      

                        if (worksheet != null)
                        {
                            // Lecture des données
                            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                            {
                               

                                // Récupération des valeurs de la ligne
                                string numeroDossierTps = worksheet.Cells[row, 1].Value?.ToString();
                                string nomOperateur = worksheet.Cells[row, 2].Value?.ToString();
                                string nomBeneficiaire = worksheet.Cells[row, 3].Value?.ToString();
                                DateTime dateDossierTps;
                                if (!DateTime.TryParse(worksheet.Cells[row, 4].Value?.ToString(), out dateDossierTps) || dateDossierTps == DateTime.MinValue)
                                    throw new ArgumentException($"La date de la ligne {row} est invalide.");
                                string codeFormulaire = worksheet.Cells[row, 5].Value?.ToString();
                                string niveauExecution = worksheet.Cells[row, 6].Value?.ToString();
                                DateTime dateRequete;
                                if (!DateTime.TryParse(worksheet.Cells[row, 7].Value?.ToString(), out dateRequete) || dateRequete == DateTime.MinValue)
                                    throw new ArgumentException($"La date de la ligne {row} est invalide.");
                                DateTime dateRetour;
                                if (!DateTime.TryParse(worksheet.Cells[row, 8].Value?.ToString(), out dateRetour) || dateRetour == DateTime.MinValue)
                                    throw new ArgumentException($"La date de la ligne {row} est invalide.");
                                string signataire = worksheet.Cells[row, 9].Value?.ToString();
                                string importationOuExportation = worksheet.Cells[row, 10].Value?.ToString();

                                // Vérification des valeurs lues
                                if (string.IsNullOrEmpty(numeroDossierTps) || string.IsNullOrEmpty(nomOperateur) || string.IsNullOrEmpty(nomBeneficiaire)
                                    || string.IsNullOrEmpty(codeFormulaire) || string.IsNullOrEmpty(niveauExecution) || string.IsNullOrEmpty(signataire)
                                    || string.IsNullOrEmpty(importationOuExportation))
                                    throw new ArgumentException($"Une ou plusieurs valeurs de la ligne {row} sont null ou vides.");

                                // Modification de la valeur de niveauExecution
                                switch (niveauExecution)
                                {
                                    case "EnCoursDouane":
                                        niveauExecution = "dlv";
                                        break;
                                    case "Amodifier":
                                        niveauExecution = "mod";
                                        break;
                                    case "Aannuler ":
                                        niveauExecution = "ann";
                                        break;
                                    case "Rejet":
                                        niveauExecution = "rej";
                                        break;
                                    case "Initialise":
                                        niveauExecution = "enc";
                                        break;
                                    default:
                                        throw new ArgumentException($"La valeur du niveau d'exécution '{niveauExecution}' de la ligne {row} n'est pas valide.");
                                }

                                // Insertion dans la base de données
                                string query = "INSERT INTO UNEPARTIEDEJOINDRE (NumeroDossierTps, CodeFormulaire, NiveauExecution, DateRequete, DateRetour, NomPole, Periode, Annee, CodeFamille, Signataire, ImportationOuExportation) " +
                                                "VALUES (@NumeroDossierTps, @CodeFormulaire, @NiveauExecution, @DateRequete, @DateRetour, @NomPole, @Periode, @Annee, @CodeFamille, @Signataire, @ImportationOuExportation)";
                                using (SqlCommand commande = new SqlCommand(query, connection))
                                {
                                    // Verifier si les parametres ne sont pas null ou vide avant l'insertion dans la base de donnée 
                                    if (string.IsNullOrEmpty(numeroDossierTps))
                                    {
                                        throw new Exception("Le numéro de dossier TPS ne peut pas être null ou vide.");
                                    }
                                    if (string.IsNullOrEmpty(nomOperateur))
                                    {
                                        throw new Exception("Le nom de l'opérateur ne peut pas être null ou vide.");
                                    }

                                    if (string.IsNullOrEmpty(nomBeneficiaire))
                                    {
                                        throw new Exception("Le nom du bénéficiaire ne peut pas être null ou vide.");
                                    }

                                    if (string.IsNullOrEmpty(codeFormulaire))
                                    {
                                        throw new Exception("Le code formulaire ne peut pas être null ou vide.");
                                    }

                                    if (string.IsNullOrEmpty(niveauExecution))
                                    {
                                        throw new Exception("Le niveau d'exécution ne peut pas être null ou vide.");
                                    }

                                    if (string.IsNullOrEmpty(signataire))
                                    {
                                        throw new Exception("Le signataire ne peut pas être null ou vide.");
                                    }

                                    if (string.IsNullOrEmpty(importationOuExportation))
                                    {
                                        throw new Exception("La valeur d'importation ou d'exportation ne peut pas être null ou vide.");
                                    }

                                    if (dateDossierTps == default(DateTime))
                                    {
                                        throw new Exception("La date de dossier TPS n'est pas valide.");
                                    }

                                    if (dateRequete == default(DateTime))
                                    {
                                        throw new Exception("La date de requête n'est pas valide.");
                                    }

                                    if (dateRetour == default(DateTime))
                                    {
                                        throw new Exception("La date de retour n'est pas valide.");
                                    }
                                    // Ajouter les paramètres
                   
                                    commande.Parameters.AddWithValue("@NumeroDossierTps", numeroDossierTps);
                                    commande.Parameters.AddWithValue("@CodeFormulaire", codeFormulaire);
                                    commande.Parameters.AddWithValue("@NiveauExecution", niveauExecution);
                                    commande.Parameters.AddWithValue("@DateRequete", dateRequete);
                                    commande.Parameters.AddWithValue("@DateRetour", dateRetour);
                                    commande.Parameters.AddWithValue("@NomPole", NomPole);
                                    commande.Parameters.AddWithValue("@Periode", Periode);
                                    commande.Parameters.AddWithValue("@Annee", Annee);
                                    commande.Parameters.AddWithValue("@CodeFamille", CodeFamille);
                                    commande.Parameters.AddWithValue("@Signataire", signataire);
                                    commande.Parameters.AddWithValue("@ImportationOuExportation", importationOuExportation);

                                    // Exécution de la commande SQL
                                    int rowsAffected = commande.ExecuteNonQuery();

                                    // Vérifier si l'insertion a réussi
                                    if (rowsAffected > 0)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        throw new Exception("L'insertion des données dans la base de données a échoué.");
                                    }

                                }

                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                // Gestion de l'exception
                throw new Exception("Une erreur est survenue lors du chargement des données : " + ex.Message);
            }
            finally
            {
                // Fermeture de la connexion à la base de données
                if (connection != null && connection.State == ConnectionState.Open)
                    connection.Close();
            }
        }



    }
}
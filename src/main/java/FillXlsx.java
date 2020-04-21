import com.spire.xls.*;

import javax.imageio.ImageIO;
import javax.rmi.CORBA.Util;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.Locale;

public class FillXlsx {

    //Private Variables
    private static File fichierSourceDureesCyclesCible;
    private static File fichierSourceDureesCyclesReference;
    private static String repertoireFichierTarget;
    private static String versionDeReference;
    private static String versionCible;

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode principale permettant de créer nos graphiques dans un fichier Synthese xlsx et de les extraire sous forme d'images PNG
    // ----------------------------------------------------------------------------------------------------------------
    public static void CreationGraphiques(File fichierSourceDureesCyclesCibleXslx, File fichierSourceDureesCyclesReferenceXslx, String repertoireFichierTargetXslx,
                                          String versionDeReferenceXslx, String versionCibleXslx) throws IOException
    {
        //Implémentation des variables
        fichierSourceDureesCyclesCible = fichierSourceDureesCyclesCibleXslx;
        fichierSourceDureesCyclesReference = fichierSourceDureesCyclesReferenceXslx;
        repertoireFichierTarget = repertoireFichierTargetXslx;
        versionDeReference = versionDeReferenceXslx;
        versionCible = versionCibleXslx;


        System.out.println("Initialisation de la création des graphiques sur document Synthese.xslx...");
        //On utilise la librairie free spire xls
        //On créer un fichier xlsx sur lequel on va écrire nos donnees et graphiques
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.getWorksheets().get(0); //Feuille 1 => 8 graphes Bars + Données Graphe Radar
        Worksheet sheet2 = workbook.getWorksheets().get(1); //Feuille 2 => Tableau Durées d'étapes
        Worksheet sheet3 = workbook.getWorksheets().get(2); //Feuille 3  => Graphe Radar


       //Le but est de récupérer les données des durées d'étapes de la version Cible et de la version de Référence
       //afin de pouvoir construire nos graphiques

        //On récupére notre fichier Durées cycle de la version Cible
        Workbook workbookFichierDureeCycleCible = new Workbook();
        //On récupére notre fichier Durées cycle de la version de référence
        Workbook workbookFichierDureeCycleRef = new Workbook();

        //on charge le fichier source Durées Cycle Cible
        workbookFichierDureeCycleCible.loadFromFile(fichierSourceDureesCyclesCible.getPath());
        //on charge le fichier source Durées Cycle Référence
        workbookFichierDureeCycleRef.loadFromFile(fichierSourceDureesCyclesReference.getPath());

        //on récupère la premiere feuille du fichier Durée cycle Cible sur laquelle se trouve les temps d'étapes de la version Cible
        Worksheet sheet0FichierDureeCycleCible = workbookFichierDureeCycleCible.getWorksheets().get(0);
        //on récupère la premiere feuille du fichier Durée cycle de Référence sur laquelle se trouve les temps d'étapes de la version de référence
        Worksheet sheet0FichierDureeCycleRef = workbookFichierDureeCycleRef.getWorksheets().get(0);


        //Création du graphique radar
        CreationGraphique1(sheet0FichierDureeCycleCible, sheet0FichierDureeCycleRef,sheet, sheet3);

        //Création du tableau récapitulatif des temps avec couleur
        CreationTableauTemps(sheet2, workbook);

        //On passe en culture info anglais pour que les décimaux avec des points puissent être pris en compte correctement
        //dans les graphiques en diagramme de bar...Par défaut on est en français...
        Locale.setDefault(new Locale("en", "US"));

        //Création des graphiques diagramme bar
        CreationGraphique2(sheet);
        CreationGraphique3(sheet);
        CreationGraphique4(sheet);
        CreationGraphique5(sheet);
        CreationGraphique6(sheet);
        CreationGraphique7(sheet);
        CreationGraphique8(sheet);

        //On sauvegarde notre fichier
        workbook.saveToFile(repertoireFichierTarget + "\\Synthese.xlsx", FileFormat.Version2016);
        System.out.println("Sauvegarde du fichier : " + repertoireFichierTarget + "\\Synthese.xlsx");
        //Puis on le réouvre par précaution afin d'être sûr d'avoir nos graphiques à disposition dans le fichier...
        workbook = new Workbook();
        workbook.loadFromFile(repertoireFichierTarget + "\\Synthese.xlsx");

        //Enfin on extrait nos graphiques du fichier Synthese.xlsx que l'on vient de créer
        //et on convertit chaque graphique en image PNG
        ExtractConvertGraphToPng(sheet,sheet2,sheet3);

    }

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode permettant d'extraire les graphes du fichier Synthese.xlsx et de les convertir en image PNG
    // ----------------------------------------------------------------------------------------------------------------
    private static void ExtractConvertGraphToPng(Worksheet sheet, Worksheet sheet2, Worksheet sheet3)
    {
        try
        {
            //Sauvegarde du graphe 1 sous forme d'image.png, la méthode ci-dessous permets d'extraire sous forme d'image une plage de cellule
            sheet3.saveToImage(repertoireFichierTarget + "\\Graphe1.png", 12,8,59,24);
            //Le problème : grosse bordures blanches autour du graphique quand on extrait l'image de la manière ci dessus
            //on est donc obligé de reconstruire une sous image sans les bordures blanches grace a la méthode getSubimage ci-dessous...hardCoreGalere oui oui...
            UtilsXlsx.ConvertGraphExcelToImgPng(repertoireFichierTarget , "Graphe1.png",  180,125,1100,860);

            sheet.saveToImage(repertoireFichierTarget + "\\Graphe2.png", 7,12,27,22); //sauvegarde du graphe2 sous forme d'image png
            UtilsXlsx.ConvertGraphExcelToImgPng(repertoireFichierTarget , "Graphe2.png", 78,105,560,397);

            sheet.saveToImage(repertoireFichierTarget + "\\Graphe3.png", 32,12,52,22);//sauvegarde du graphe3 sous forme d'image png
            UtilsXlsx.ConvertGraphExcelToImgPng(repertoireFichierTarget , "Graphe3.png", 78,105,560,397);


            sheet.saveToImage(repertoireFichierTarget + "\\Graphe4.png", 56,12,73,22);//sauvegarde du graphe4 sous forme d'image png
            UtilsXlsx.ConvertGraphExcelToImgPng(repertoireFichierTarget , "Graphe4.png", 78,105,320,340);


            sheet.saveToImage(repertoireFichierTarget + "\\Graphe5.png", 77,12,95,22);//sauvegarde du graphe5 sous forme d'image png
            UtilsXlsx.ConvertGraphExcelToImgPng(repertoireFichierTarget , "Graphe5.png", 78,105,320,340);


            sheet.saveToImage(repertoireFichierTarget + "\\Graphe6.png", 101,12,119,22);//sauvegarde du graphe6 sous forme d'image png
            UtilsXlsx.ConvertGraphExcelToImgPng(repertoireFichierTarget , "Graphe6.png", 78,105,320,340);


            sheet.saveToImage(repertoireFichierTarget + "\\Graphe7.png", 125,12,146,22);//sauvegarde du graphe7 sous forme d'image png
            UtilsXlsx.ConvertGraphExcelToImgPng(repertoireFichierTarget , "Graphe7.png", 78,105,480,415);


            sheet.saveToImage(repertoireFichierTarget + "\\Graphe8.png", 150,12,172,22);//sauvegarde du graphe8 sous forme d'image png
            UtilsXlsx.ConvertGraphExcelToImgPng(repertoireFichierTarget , "Graphe8.png", 78,105,480,415);


            sheet2.saveToImage(repertoireFichierTarget + "\\Synthese.png", 1,2,11,5);//sauvegarde du tableau de temps
            UtilsXlsx.ConvertGraphExcelToImgPng(repertoireFichierTarget , "Synthese.png", 77,106,958,234);


            //A noter que l'image de la légende version Cible/version Référence est extraite directement de l'image du Graphe2.png
            BufferedImage legende = ImageIO.read(new File(repertoireFichierTarget + "\\Graphe2.png"));
            legende = legende.getSubimage(3,15,550,35);
            //Puis on construit notre nouvelle image redimensionnée...
            ImageIO.write(legende, "png", new File(repertoireFichierTarget + "\\LegendeVersions.png"));
            System.out.println(Utils.ANSI_BLUE + "Création LegendeVersions.png OK" + Utils.ANSI_RESET);

        } catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Création du graphique 1 Radar sur la premiere feuille
    // ----------------------------------------------------------------------------------------------------------------
    private static void CreationGraphique1(Worksheet sheet0FichierDureeCycleCible, Worksheet sheet0FichierDureeCycleRef, Worksheet sheet, Worksheet sheet3)
    {
        sheet.getCellRange("A1").setText("Etapes (nbre heures)/Versions");
        sheet.getCellRange("A2").setText("Retour Solde");
        sheet.getCellRange("A3").setText("Calcul CMC");
        sheet.getCellRange("A4").setText("Editions CMC");
        sheet.getCellRange("A5").setText("TPC");
        sheet.getCellRange("A6").setText("Calcul CNQ");
        sheet.getCellRange("A7").setText("Editions CNQ");
        sheet.getCellRange("A8").setText("Duplicata BMS Jasper");
        sheet.getCellRange("A9").setText("Flux de masse");
        sheet.getCellRange("A10").setText("Flux de valorisation");
        sheet.getCellRange("A11").setText("Total");

        sheet.getCellRange("B1").setText("Durée Etape (%) " + versionDeReference);
        sheet.getCellRange("C1").setText("Durée Etape (%) " + versionCible);

        sheet.getCellRange("D1").setText("Tps Etape (heure) " + versionDeReference);
        sheet.getCellRange("E1").setText("Tps Etape (heure) " + versionCible);



        //On remplit notre troisieme colonne avec les temps d'étape de la version de référence
        RemplissageRadarColonneTempsEtape(sheet0FichierDureeCycleRef, sheet, "D");

        //On remplit notre quatrieme colonne avec les temps d'étape de la version cible:
        RemplissageRadarColonneTempsEtape(sheet0FichierDureeCycleCible, sheet, "E");

        //On remplit notre premiere colonne avec les pourcentage de temps d'étape de la version de référence
        RemplissageRadarColonnePourcentageTempsEtape(sheet, "B", "D");

        //On remplit notre deuxieme colonne avec les pourcentage de temps d'étape de la version cible
        RemplissageRadarColonnePourcentageTempsEtape(sheet, "C", "E");

        Chart chart = sheet3.getCharts().add(ExcelChartType.Radar);
        chart.setChartTitle(" ");
        chart.setDataRange(sheet.getCellRange("Sheet1!A1:C10"));
        chart.setValueAxisTitle("test");
        chart.setLeftColumn(1);
        chart.setRightColumn(32);
        chart.setTopRow(12);
        chart.setBottomRow(60);
        chart.setSeriesDataFromRange(false);
        chart.setPlotVisibleOnly(false);
        chart.getLegend().setPosition(LegendPositionType.Top);
        chart.setName("Graphique1");
        System.out.println(Utils.ANSI_BLUE + "Fin création Graphique 1 OK" + Utils.ANSI_RESET);

    }

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode permettant de remplir la colonne Temps Etape pour le graphique Radar
    // ----------------------------------------------------------------------------------------------------------------
    private static void RemplissageRadarColonneTempsEtape(Worksheet sheet0FichierDureeCycle, Worksheet sheet, String lettreColonne)
    {
        String valueC5 = sheet0FichierDureeCycle.getCellRange("C5").getEnvalutedValue(); //on récupère la valeur
        sheet.getCellRange(lettreColonne +"2").setNumberValue(UtilsXlsx.ConvertohhmmssToNbreHeureDecimal(valueC5));

        String valueC6 = sheet0FichierDureeCycle.getCellRange("C6").getEnvalutedValue(); //on récupère la valeur
        sheet.getCellRange(lettreColonne +"3").setNumberValue(UtilsXlsx.ConvertohhmmssToNbreHeureDecimal(valueC6));

        String valueC7 = sheet0FichierDureeCycle.getCellRange("C7").getEnvalutedValue(); //on récupère la valeur
        sheet.getCellRange(lettreColonne +"4").setNumberValue(UtilsXlsx.ConvertohhmmssToNbreHeureDecimal(valueC7));

        String valueC8 = sheet0FichierDureeCycle.getCellRange("C8").getEnvalutedValue(); //on récupère la valeur
        sheet.getCellRange(lettreColonne +"5").setNumberValue(UtilsXlsx.ConvertohhmmssToNbreHeureDecimal(valueC8));

        String valueC9 = sheet0FichierDureeCycle.getCellRange("C9").getEnvalutedValue(); //on récupère la valeur
        sheet.getCellRange(lettreColonne +"6").setNumberValue(UtilsXlsx.ConvertohhmmssToNbreHeureDecimal(valueC9));

        String valueC10 = sheet0FichierDureeCycle.getCellRange("C10").getEnvalutedValue(); //on récupère la valeur
        sheet.getCellRange(lettreColonne +"7").setNumberValue(UtilsXlsx.ConvertohhmmssToNbreHeureDecimal(valueC10));

        String valueC11 = sheet0FichierDureeCycle.getCellRange("C11").getEnvalutedValue(); //on récupère la valeur
        sheet.getCellRange(lettreColonne +"8").setNumberValue(UtilsXlsx.ConvertohhmmssToNbreHeureDecimal(valueC11));

        String valueC12 = sheet0FichierDureeCycle.getCellRange("C12").getEnvalutedValue(); //on récupère la valeur
        sheet.getCellRange(lettreColonne +"9").setNumberValue(UtilsXlsx.ConvertohhmmssToNbreHeureDecimal(valueC12));

        String valueC13 = sheet0FichierDureeCycle.getCellRange("C13").getEnvalutedValue(); //on récupère la valeur
        sheet.getCellRange(lettreColonne +"10").setNumberValue(UtilsXlsx.ConvertohhmmssToNbreHeureDecimal(valueC13));
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode permettant de remplir la colonne Pourcentage Temps Etape pour le graphique Radar
    // ----------------------------------------------------------------------------------------------------------------
    private static void RemplissageRadarColonnePourcentageTempsEtape( Worksheet sheet, String lettreColonnePourcentage, String lettreColonneTemps)
    {
        Double value1 = Double.parseDouble(sheet.getCellRange(lettreColonneTemps + "2").getValue());
        Double value2 =  Double.parseDouble(sheet.getCellRange(lettreColonneTemps + "3").getValue());
        Double value3 =  Double.parseDouble(sheet.getCellRange(lettreColonneTemps + "4").getValue());
        Double value4 =  Double.parseDouble(sheet.getCellRange(lettreColonneTemps + "5").getValue());
        Double value5 =  Double.parseDouble(sheet.getCellRange(lettreColonneTemps + "6").getValue());
        Double value6 =  Double.parseDouble(sheet.getCellRange(lettreColonneTemps + "7").getValue());
        Double value7 =  Double.parseDouble(sheet.getCellRange(lettreColonneTemps + "8").getValue());
        Double value8 =  Double.parseDouble(sheet.getCellRange(lettreColonneTemps + "9").getValue());
        Double value9 =  Double.parseDouble(sheet.getCellRange(lettreColonneTemps + "10").getValue());
        Double total = value1 + value2 + value3 + value4 + value5 + value6 + value7 + value8 + value9; //récupération du temps total
        //On remplit la cellule du total du temps
        sheet.getCellRange(lettreColonneTemps +"11").setValue(String.valueOf(total));
        //Ensuite on établit le pourcentage par rapport au temps total pour chaque étape
        sheet.getCellRange(lettreColonnePourcentage +"2").setNumberValue(UtilsXlsx.ConvertStringEvaluatedValuePourcentToDouble(String.valueOf(value1/total)));
        sheet.getCellRange(lettreColonnePourcentage +"3").setNumberValue(UtilsXlsx.ConvertStringEvaluatedValuePourcentToDouble(String.valueOf(value2/total)));
        sheet.getCellRange(lettreColonnePourcentage +"4").setNumberValue(UtilsXlsx.ConvertStringEvaluatedValuePourcentToDouble(String.valueOf(value3/total)));
        sheet.getCellRange(lettreColonnePourcentage +"5").setNumberValue(UtilsXlsx.ConvertStringEvaluatedValuePourcentToDouble(String.valueOf(value4/total)));
        sheet.getCellRange(lettreColonnePourcentage +"6").setNumberValue(UtilsXlsx.ConvertStringEvaluatedValuePourcentToDouble(String.valueOf(value5/total)));
        sheet.getCellRange(lettreColonnePourcentage +"7").setNumberValue(UtilsXlsx.ConvertStringEvaluatedValuePourcentToDouble(String.valueOf(value6/total)));
        sheet.getCellRange(lettreColonnePourcentage +"8").setNumberValue(UtilsXlsx.ConvertStringEvaluatedValuePourcentToDouble(String.valueOf(value7/total)));
        sheet.getCellRange(lettreColonnePourcentage +"9").setNumberValue(UtilsXlsx.ConvertStringEvaluatedValuePourcentToDouble(String.valueOf(value8/total)));
        sheet.getCellRange(lettreColonnePourcentage +"10").setNumberValue(UtilsXlsx.ConvertStringEvaluatedValuePourcentToDouble(String.valueOf(value9/total)));
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode permettant de créer le Tableau Temps d'étape pour la Slide 3
    // ----------------------------------------------------------------------------------------------------------------
    private static void CreationTableauTemps(Worksheet sheet2, Workbook workbook)
    {
        //On va mettre ce tableau sur la sheet2 du fichier excel car on va manipuler les tailles des colonnes (pour l'extraire de manière propre sous format png)
        // et cela impacte sur la taille des graphiques déjà présents sur la sheet1...

        //Libellés des colonnes
        sheet2.getCellRange("B2").setText("Retour Solde");
        sheet2.getCellRange("B3").setText("Calcul CMC");
        sheet2.getCellRange("B4").setText("Editions CMC");
        sheet2.getCellRange("B5").setText("TPC");
        sheet2.getCellRange("B6").setText("Calcul CNQ");
        sheet2.getCellRange("B7").setText("Editions CNQ");
        sheet2.getCellRange("B8").setText("Duplicata BMS Jasper");
        sheet2.getCellRange("B9").setText("Flux de masse");
        sheet2.getCellRange("B10").setText("Flux de valorisation");
        sheet2.getCellRange("B11").setText("Total");
        CellRange crLibelles = sheet2.getCellRange("B2:B11");
        crLibelles.getStyle().setHorizontalAlignment(HorizontalAlignType.Center);
        crLibelles.getStyle().setColor(UtilsXlsx.convertRgbToHsb(225, 239, 245)); //background bleu clair

        //Titres principaux
        sheet2.getCellRange("B1").setText("Etapes/Temps(heures)");
        sheet2.getCellRange("C1").setText(versionDeReference);
        sheet2.getCellRange("D1").setText(versionCible);
        //Mise en forme des titres principaux
        CellStyle styleTitre = workbook.getStyles().addStyle("styleTitreTemps"); //pour le fun j'ai crée un style
        styleTitre.setColor(UtilsXlsx.convertRgbToHsb(141, 200, 227)); //background bleu
        styleTitre.getFont().setColor(Color.white); //couleur texte
        styleTitre.getFont().isBold(true); //texte en gras
        styleTitre.setVerticalAlignment(VerticalAlignType.Center);
        styleTitre.setHorizontalAlignment(HorizontalAlignType.Center.Center);
        CellRange crTitres = sheet2.getCellRange("B1:D1");
        crTitres.setStyle(styleTitre);

        sheet2.setRowHeight(1,30); //hauteur de la cellule Etapes/temps
        sheet2.setColumnWidth(2, 26); //largeur de la colonne 2
        sheet2.setColumnWidth(3, 40); //largeur de la colonne 3
        sheet2.setColumnWidth(4, 40); //largeur de la colonne 4



        //Données des temps (on les reprends sur le tableau initial
        sheet2.getCellRange("C2").setFormula("=Sheet1!D2");
        sheet2.getCellRange("C3").setFormula("=Sheet1!D3");
        sheet2.getCellRange("C4").setFormula("=Sheet1!D4");
        sheet2.getCellRange("C5").setFormula("=Sheet1!D5");
        sheet2.getCellRange("C6").setFormula("=Sheet1!D6");
        sheet2.getCellRange("C7").setFormula("=Sheet1!D7");
        sheet2.getCellRange("C8").setFormula("=Sheet1!D8");
        sheet2.getCellRange("C9").setFormula("=Sheet1!D9");
        sheet2.getCellRange("C10").setFormula("=Sheet1!D10");
        sheet2.getCellRange("C11").setFormula("=Sheet1!D11");

        sheet2.getCellRange("D2").setFormula("=Sheet1!E2");
        sheet2.getCellRange("D3").setFormula("=Sheet1!E3");
        sheet2.getCellRange("D4").setFormula("=Sheet1!E4");
        sheet2.getCellRange("D5").setFormula("=Sheet1!E5");
        sheet2.getCellRange("D6").setFormula("=Sheet1!E6");
        sheet2.getCellRange("D7").setFormula("=Sheet1!E7");
        sheet2.getCellRange("D8").setFormula("=Sheet1!E8");
        sheet2.getCellRange("D9").setFormula("=Sheet1!E9");
        sheet2.getCellRange("D10").setFormula("=Sheet1!E10");
        sheet2.getCellRange("D11").setFormula("=Sheet1!E11");

        //Mise en forme des données
        CellRange crDonnees = sheet2.getCellRange("C2:D11");
        crDonnees.getStyle().setHorizontalAlignment(HorizontalAlignType.Center);
        crDonnees.getStyle().getFont().setSize(8);
        sheet2.getCellRange("D2:D11").getStyle().getFont().isBold(true);

        //Traitement couleur sur la colonne des durées de la version cible
        traitementCouleurTpsVersionCible(sheet2);

        //Mise en forme globale:
        CellRange crAll = sheet2.getCellRange("B1:D11");
        crAll.getBorders().getByBordersLineType(BordersLineType.EdgeTop).setLineStyle(LineStyleType.Thin); //Bordure haute
        crAll.getBorders().getByBordersLineType(BordersLineType.EdgeBottom).setLineStyle(LineStyleType.Thin); //bordure basse
        crAll.getBorders().getByBordersLineType(BordersLineType.EdgeLeft).setLineStyle(LineStyleType.Thin);// Bordure gauche
        crAll.getBorders().getByBordersLineType(BordersLineType.EdgeRight).setLineStyle(LineStyleType.Thin); //Bordure Droite


    }

    // ----------------------------------------------------------------------------------------------------------------
    //Méthode permettant de colorer le temps d'étape selon qu'il soit inférieur ou supérieur à celui de référence
    // (dans le tableau de durées d'étape de la slide 3)
    // ----------------------------------------------------------------------------------------------------------------
    private static void traitementCouleurTpsVersionCible(Worksheet sheet)
    {
        //on affiche en vert si le temps est inférieur à celui de la version de référence
        //ou en rouge si le temps est supérieur à celui de la version de référence
        for (int i = 2; i < 12; i++)
        {
            double valueVersionRef = Double.parseDouble(sheet.getCellRange("C" + i).getEnvalutedValue());
            double valueVersionCible = Double.parseDouble(sheet.getCellRange("D" + i).getEnvalutedValue());
            if(valueVersionCible <= valueVersionRef)
            {
                sheet.getCellRange("D"+i).getStyle().getFont().setColor(UtilsXlsx.convertRgbToHsb(39,174,96)); //vert
            }
            else {
                sheet.getCellRange("D"+i).getStyle().getFont().setColor(UtilsXlsx.convertRgbToHsb(192,57,43)); //rouge
            }
        }
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Création du graphique 2 Durée Editique CMC CNQ
    // ----------------------------------------------------------------------------------------------------------------
    private static void CreationGraphique2(Worksheet sheet)
    {
        //Attention ne pas se fier à la disposition des titres du tableau
        sheet.getCellRange("L5").setText("Durée Editique CMC");
        sheet.getCellRange("L6").setText("Durée Editique CNQ");
        sheet.getCellRange("M4").setText(versionDeReference);
        sheet.getCellRange("N4").setText(versionCible);

        //Pour les valeurs pour les graphiques, ne pas se fier aux titres
        //-- en effet M5 et M6 correspondent respectivement à la durée éditique CMC de la version de ref
        //et à la durée éditique CNQ de la version de ref
        //-- en effet N5 et N6 correspondent respectivement à la durée éditique CMC de la version de cible
        //et à la durée éditique CNQ de la version de cible
        String valueD4 = sheet.getCellRange("D4").getValue();
        sheet.getCellRange("M5").setValue(valueD4);

        String valueD7 = sheet.getCellRange("D7").getValue();
        sheet.getCellRange("M6").setValue(valueD7);

        String valueE4 = sheet.getCellRange("E4").getValue();
        sheet.getCellRange("N5").setValue(valueE4);

        String valueE7 = sheet.getCellRange("E7").getValue();
        sheet.getCellRange("N6").setValue(valueE7);



        Chart chart = sheet.getCharts().add(ExcelChartType.ColumnClustered);
        chart.setChartTitle("");//initialisé comme ceci car si je ne le mets pas il mets un titre chinois au graphique...
        chart.getChartTitleArea().setSize(8);
        chart.setDataRange(sheet.getCellRange("L4:N6"));
        chart.setValueAxisTitle("Heures");
        //positions du graphe...
        chart.setTopRow(7);
        chart.setLeftColumn(12);
        chart.setRightColumn(19);
        chart.setBottomRow(28);

        chart.setSeriesDataFromRange(false);
        chart.setPlotVisibleOnly(false);
        chart.getLegend().setPosition(LegendPositionType.Top);
        chart.setName("Graphique2");
        System.out.println(Utils.ANSI_BLUE +  "Fin création Graphique 2 OK" + Utils.ANSI_RESET);
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Création du graphique 3 Durée CMC CNQ
    // ----------------------------------------------------------------------------------------------------------------
    private static void CreationGraphique3(Worksheet sheet)
    {
        //Attention ne pas se fier à la disposition des titres du tableau
        //voir juste après pourquoi
        sheet.getCellRange("L30").setText("Durée CMC");
        sheet.getCellRange("L31").setText("Durée CNQ");
        sheet.getCellRange("M29").setText(versionDeReference);
        sheet.getCellRange("N29").setText(versionCible);

        //Pour les valeurs pour les graphiques, ne pas se fier aux titres
        //-- en effet M30 et M31correspondent respectivement à la durée  CMC de la version de ref
        //et à la durée  CNQ de la version de ref
        //-- en effet N30 et N31 correspondent respectivement à la durée  CMC de la version de cible
        //et à la durée  CNQ de la version de cible
        String valueD3 = sheet.getCellRange("D3").getValue();
        sheet.getCellRange("M30").setValue(valueD3);

        String valueD6 = sheet.getCellRange("D6").getValue();
        sheet.getCellRange("M31").setValue(valueD6);

        String valueE3 = sheet.getCellRange("E3").getValue();
        sheet.getCellRange("N30").setValue(valueE3);

        String valueE6 = sheet.getCellRange("E6").getValue();
        sheet.getCellRange("N31").setValue(valueE6);



        Chart chart = sheet.getCharts().add(ExcelChartType.ColumnClustered);
        chart.setChartTitle(""); //initialisé comme ceci car si je ne le mets pas il mets un titre chinois au graphique...
        chart.getChartTitleArea().setSize(8);
        chart.setDataRange(sheet.getCellRange("L29:N31"));
        chart.setValueAxisTitle("Heures");
        //positions du graphe...
        chart.setTopRow(32);
        chart.setBottomRow(53);
        chart.setLeftColumn(12);
        chart.setRightColumn(19);


        chart.setSeriesDataFromRange(false);
        chart.setPlotVisibleOnly(false);
        chart.getLegend().setPosition(LegendPositionType.Top);
        chart.setName("Graphique3");
        System.out.println(Utils.ANSI_BLUE +  "Fin création Graphique 3 OK" + Utils.ANSI_RESET);
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Création du graphique 4 Durée Flux de masse
    // ----------------------------------------------------------------------------------------------------------------
    private static void CreationGraphique4(Worksheet sheet)
    {
        //Attention ne pas se fier à la disposition des titres du tableau
        //voir juste après pourquoi
        sheet.getCellRange("L55").setText("Flux de masse");
        sheet.getCellRange("M54").setText(versionDeReference);
        sheet.getCellRange("N54").setText(versionCible);

        //Pour les valeurs pour les graphiques, ne pas se fier aux titres
        //-- en effet M55 et N55 correspondent respectivement à la durée  flux de masse de la version de ref
        //et à la durée  flux de masse de la version cible
        String valueD9 = sheet.getCellRange("D9").getValue();
        sheet.getCellRange("M55").setValue(valueD9);

        String valueE9 = sheet.getCellRange("E9").getValue();
        sheet.getCellRange("N55").setValue(valueE9);




        Chart chart = sheet.getCharts().add(ExcelChartType.ColumnClustered);
        chart.setChartTitle(""); //initialisé comme ceci car si je ne le mets pas il mets un titre chinois au graphique...
        chart.getChartTitleArea().setSize(8);
        chart.setDataRange(sheet.getCellRange("L54:N55"));
        chart.setValueAxisTitle("Heures");
        //positions du graphe...
        chart.setTopRow(56);
        chart.setBottomRow(74);
        chart.setLeftColumn(12);
        chart.setRightColumn(16);


        chart.setSeriesDataFromRange(false);
        chart.setPlotVisibleOnly(false);
        chart.getLegend().setPosition(LegendPositionType.Top);
        chart.getLegend().delete();
        chart.setName("Graphique4");
        System.out.println(Utils.ANSI_BLUE +  "Fin création Graphique 4 OK" + Utils.ANSI_RESET);
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Création du graphique 5 Durée duplicata BMS Jasper
    // ----------------------------------------------------------------------------------------------------------------
    private static void CreationGraphique5(Worksheet sheet)
    {
        //Attention ne pas se fier à la disposition des titres du tableau
        //voir juste après pourquoi
        sheet.getCellRange("L76").setText("Duplicatas BMS Jasper");
        sheet.getCellRange("M75").setText(versionDeReference);
        sheet.getCellRange("N75").setText(versionCible);

        //Pour les valeurs pour les graphiques, ne pas se fier aux titres
        //-- en effet M76 et N76 correspondent respectivement à la durée  BMS Jasper de la version de ref
        //et à la durée  BMS Jasper de la version cible
        String valueD8 = sheet.getCellRange("D8").getValue();
        sheet.getCellRange("M76").setValue(valueD8);

        String valueE8 = sheet.getCellRange("E8").getValue();
        sheet.getCellRange("N76").setValue(valueE8);




        Chart chart = sheet.getCharts().add(ExcelChartType.ColumnClustered);
        chart.setChartTitle(""); //initialisé comme ceci car si je ne le mets pas il mets un titre chinois au graphique...
        chart.getChartTitleArea().setSize(8);
        chart.setDataRange(sheet.getCellRange("L75:N76"));
        chart.setValueAxisTitle("Heures");
        //positions du graphe...
        chart.setTopRow(77);
        chart.setBottomRow(95);
        chart.setLeftColumn(12);
        chart.setRightColumn(16);


        chart.setSeriesDataFromRange(false);
        chart.setPlotVisibleOnly(false);
        chart.getLegend().setPosition(LegendPositionType.Top);
        chart.getLegend().delete();
        chart.setName("Graphique5");
        System.out.println(Utils.ANSI_BLUE +  "Fin création Graphique 5 OK" + Utils.ANSI_RESET);
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Création du graphique 6 Durée Flux de Valorisation
    // ----------------------------------------------------------------------------------------------------------------
    private static void CreationGraphique6(Worksheet sheet)
    {
        //Attention ne pas se fier à la disposition des titres du tableau
        //voir juste après pourquoi
        sheet.getCellRange("L100").setText("Flux de VALO");
        sheet.getCellRange("M99").setText(versionDeReference);
        sheet.getCellRange("N99").setText(versionCible);

        //Pour les valeurs pour les graphiques, ne pas se fier aux titres
        //-- en effet M100 et N100 correspondent respectivement à la durée  BMS Jasper de la version de ref
        //et à la durée  BMS Jasper de la version cible
        String valueD10 = sheet.getCellRange("D10").getValue();
        sheet.getCellRange("M100").setValue(valueD10);

        String valueE10 = sheet.getCellRange("E10").getValue();
        sheet.getCellRange("N100").setValue(valueE10);

        Chart chart = sheet.getCharts().add(ExcelChartType.ColumnClustered);
        chart.setChartTitle(""); //initialisé comme ceci car si je ne le mets pas il mets un titre chinois au graphique...
        chart.getChartTitleArea().setSize(8);
        chart.setDataRange(sheet.getCellRange("L99:N100"));
        chart.setValueAxisTitle("Heures");
        //positions du graphe...
        chart.setTopRow(101);
        chart.setBottomRow(119);
        chart.setLeftColumn(12);
        chart.setRightColumn(16);


        chart.setSeriesDataFromRange(false);
        chart.setPlotVisibleOnly(false);
        chart.getLegend().delete();
        chart.getLegend().setPosition(LegendPositionType.Top);
        chart.setName("Graphique6");
        System.out.println(Utils.ANSI_BLUE +  "Fin création Graphique 6 OK" + Utils.ANSI_RESET);
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Création du graphique 7 Durée TPC
    // ----------------------------------------------------------------------------------------------------------------
    private static void CreationGraphique7(Worksheet sheet)
    {
        //Attention ne pas se fier à la disposition des titres du tableau
        //voir juste après pourquoi
        sheet.getCellRange("L124").setText("TPC");
        sheet.getCellRange("M123").setText(versionDeReference);
        sheet.getCellRange("N123").setText(versionCible);

        //Pour les valeurs pour les graphiques, ne pas se fier aux titres
        //-- en effet M124 et N124 correspondent respectivement à la durée  TPC de la version de ref
        //et à la durée  TPC de la version cible
        String valueD5 = sheet.getCellRange("D5").getValue();
        sheet.getCellRange("M124").setValue(valueD5);

        String valueE5 = sheet.getCellRange("E5").getValue();
        sheet.getCellRange("N124").setValue(valueE5);

        Chart chart = sheet.getCharts().add(ExcelChartType.ColumnClustered);
        chart.setChartTitle(""); //initialisé comme ceci car si je ne le mets pas il mets un titre chinois au graphique...
        chart.getChartTitleArea().setSize(8);
        chart.setDataRange(sheet.getCellRange("L123:N124"));
        chart.setValueAxisTitle("Heures");
        //positions du graphe...
        chart.setTopRow(125);
        chart.setBottomRow(147);
        chart.setLeftColumn(12);
        chart.setRightColumn(18);


        chart.setSeriesDataFromRange(false);
        chart.setPlotVisibleOnly(false);
        chart.getLegend().setPosition(LegendPositionType.Top);
        chart.setName("Graphique7");
        chart.getLegend().delete();
        System.out.println( Utils.ANSI_BLUE + "Fin création Graphique 7 OK" + Utils.ANSI_RESET);
    }

    // ----------------------------------------------------------------------------------------------------------------
    //Création du graphique 8 Durée TPC
    // ----------------------------------------------------------------------------------------------------------------
    private static void CreationGraphique8(Worksheet sheet)
    {
        //Attention ne pas se fier à la disposition des titres du tableau
        //voir juste après pourquoi
        sheet.getCellRange("L149").setText("Calcul CNQ");
        sheet.getCellRange("M148").setText(versionDeReference);
        sheet.getCellRange("N148").setText(versionCible);

        //Pour les valeurs pour les graphiques, ne pas se fier aux titres
        //-- en effet M148 et N148 correspondent respectivement à la durée  du calcul CNQ de la version de ref
        //et à la durée  du calcul CNQ de la version cible
        String valueD6 = sheet.getCellRange("D6").getValue();
        sheet.getCellRange("M149").setValue(valueD6);

        String valueE6 = sheet.getCellRange("E6").getValue();
        sheet.getCellRange("N149").setValue(valueE6);

        Chart chart = sheet.getCharts().add(ExcelChartType.ColumnClustered);
        chart.setChartTitle(""); //initialisé comme ceci car si je ne le mets pas il mets un titre chinois au graphique...
        chart.getChartTitleArea().setSize(8);
        chart.setDataRange(sheet.getCellRange("L148:N149"));
        chart.setValueAxisTitle("Heures");
        //positions du graphe...
        chart.setTopRow(150);
        chart.setBottomRow(172);
        chart.setLeftColumn(12);
        chart.setRightColumn(18);


        chart.setSeriesDataFromRange(false);
        chart.setPlotVisibleOnly(false);
        chart.getLegend().delete();
        chart.getLegend().setPosition(LegendPositionType.Top);
        chart.setName("Graphique8");
        System.out.println( Utils.ANSI_BLUE + "Fin création Graphique 8 OK" + Utils.ANSI_RESET);
    }

}

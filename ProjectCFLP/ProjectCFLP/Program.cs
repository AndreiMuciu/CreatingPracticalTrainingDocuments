using System;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
using Xceed.Document.NET;
using Xceed.Words.NET;

class Program
{
    public class Student
    {
        public string StudentName { get; set; }
        public string StudentSurname { get; set; }
        public int StudentBornYear { get; set; }
        public string StudentEmail { get; set; }
        public string StudentPhoneNumber { get; set; }
        public string StudentCNP { get; set; }
        public string StudentAddress { get; set; }
        public string StudentCitizenship { get; set; }

        public string GetFullName()
        {
            return $"{StudentName} {StudentSurname}";
        }

        public int GetAge()
        {
            return DateTime.Now.Year - StudentBornYear;
        }

        public override string ToString()
        {
            return $"{StudentName} {StudentSurname} - {this.GetAge()} years old - {StudentEmail} - {StudentPhoneNumber} - {StudentCNP}";
        }
    }

    public class Company
    {
        public string CompanyName { get; set; }
        public string CompanyAddress { get; set; }
        public string CompanyOfficialPerson { get; set; }
        public string CompanyPhoneNumber { get; set; }
        public string CompanyEmail { get; set; }
        public string CompanyFiscalCode { get; set; }
        public string CompanyRegistrationNumber { get; set; }
    }

    public class Tutor
    {
        public string TuturName { get; set; }
        public string TutorSurname { get; set; }
        public string TutorEmail { get; set; }
        public string TutorPhoneNumber { get; set; }
        public string TutorPosition { get; set; }

        public string getFullName()
        {
            return $"{TuturName} {TutorSurname}";
        }
    }

    public class TeachingStaffSupervisor
    {
        public string TSSName { get; set; }
        public string TSSEmail { get; set; }
        public string TSSPhoneNumber { get; set; }
        public string TSSPosition { get; set; }
    }

    public class PracticeInformation
    {
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public int CreditsNumber { get; set; }
        public int HoursNumber { get; set; }
    }

    public class PracticeConvention
    {
        private static int _nextId = 1; // Static counter for generating unique IDs
        public int Id { get; private set; } // Unique ID
        public Student Stud { get; set; }
        public Company Comp { get; set; }
        public Tutor Tut { get; set; }
        public TeachingStaffSupervisor Tss { get; set; }
        public PracticeInformation Information { get; set; }

        public PracticeConvention(Student student, Company company, Tutor tutor, TeachingStaffSupervisor tss, PracticeInformation info)
        {
            Id = _nextId++;
            Stud = student;
            Comp = company;
            Tut = tutor;
            Tss = tss;
            Information = info;

        }
        public override string ToString()
        {
            return Stud.StudentName + " - " + Comp.CompanyName;
        }

        public void WriteConventionToWordEN()
        {
            using (var document = DocX.Create(Stud.StudentName + Stud.StudentSurname + "-" + Comp.CompanyName + "ConventieEN.docx"))
            {
                document.InsertParagraph("FRAMEWORK AGREEMENT REGARDING THE COMPLETION OF THE PRACTICAL TRAINING STIPULATED IN THE UNDERGRADUATE CURRICULUM (Bachelor and Master’s Levels)")
                       .FontSize(14)
                       .Bold()
                       .Alignment = Alignment.center;

                document.InsertParagraph("\nThis framework agreement shall be closed between:")
                       .FontSize(12)
                       .SpacingAfter(10);

                document.InsertParagraph("Politehnica University of Timișoara (hereinafter referred to as practical training organiser), represented by Rector Assoc. Prof. Florin DRĂGAN, PhD Eng, the organiser’s address being: TIMIȘOARA, postal code 300 006, Piața Victoriei, nr. 2, telephone: 0256.40301, e-mail: rector@upt.ro; unique registration code: 4269282, hereinafter referred to as the Practical Training Organizer.")
                       .FontSize(11)
                       .SpacingAfter(10);

                document.InsertParagraph($"{Comp.CompanyName}(hereinafter referred to as practical training partner), represented by (name and capacity) {Comp.CompanyOfficialPerson}, the practical training partner's address being: {Comp.CompanyAddress} the address where practical training shall take place: {Comp.CompanyAddress} email: {Comp.CompanyEmail}, telephone {Comp.CompanyPhoneNumber}; hereinafter referred to as the Practical Training Partner")
                       .FontSize(11)
                       .Italic()
                       .SpacingAfter(10);

                document.InsertParagraph($"Student {Stud.StudentName} (hereinafter referred to as trainee), personal identification number {Stud.StudentCNP} date of birth {Stud.StudentBornYear}, place of birth {Stud.StudentAddress}, citizenship {Stud.StudentCitizenship} , residence address {Stud.StudentAddress} email: {Stud.StudentEmail}, telephone: {Stud.StudentPhoneNumber} hereinafter referred to as the Practical Trainee")
                       .FontSize(11)
                       .SpacingAfter(15);

                InsertArticle(document, "Art 1. Object of the framework agreement",
                    "This framework agreement sets up the framework under which the practical training is organised and carried out to strengthen the theoretical knowledge and to develop the training of practical skills, to apply them in accordance with the specialization for which the practicing student is trained.");

                InsertArticle(document, "Art 2. Status of the trainee",
                    "During the whole practical training, the trainee shall maintain his/her role as student at Politehnica University of Timisoara.");

                InsertArticle(document, "Art 3. Duration of the practical training",
                    $"The practical training as described in the Academic Curricula consists of {Information.HoursNumber}[hrs.]\n\n" +
                    $"The practical training period is in accordance with the structure of the current academic year, from {Information.StartDate} (day/month/year) until {Information.EndDate} (day/month/year).");

                InsertArticle(document, "Art 4. Payment and social obligations",
                    "Practical training stage (tick the appropriate situation):\n\n" +
                    "☐ is performed within an employment contract, the two partners being able to benefit from the provisions of Law no. 72/2007 on stimulating the employment of pupils and students;\n\n" +
                    "☐ is not performed under an employment contract;\n\n" +
                    "☐ is carried out in the framework of a project financed by the European Social Fund;\n\n" +
                    "☐ is carried out within the project ......");

                InsertArticle(document, "Art 5. Responsibilities of the trainee",
                    "During the practical training, the trainee must comply with the stipulated work schedule and carry out the activities specified by the tutor in accordance with the practical training program adhering to the legal framework concerning their volume and difficulty.\n\n" +
                    "The trainee shall have the obligation to comply with the occupational health and safety regulations.");

                InsertArticle(document, "Art 6. Responsibilities of the practical training partner",
                    "The training partner shall appoint a practical training supervisor, selected from its employees, whose obligations are mentioned in the practical training portfolio, an integral part of the framework convention.");

                InsertArticle(document, "Art 7. Obligations of the practical training organiser",
                    "The practical training organiser shall appoint a supervising member of the teaching staff, in charge with planning, organising and supervising the completion of the practical training.");

                InsertArticle(document, "Art 8. Persons appointed by the practical training organiser and by the practical training partner",
                    "Practical training supervisor (person guiding the trainee, appointed by the practical training partner):\n\n" +
                    $"Mr./Ms {Tss.TSSName}\nRole {Tss.TSSPosition} Telephone {Tss.TSSPhoneNumber}");

                InsertArticle(document, "Art 9. Assessment of the practical training through credit transfer",
                    "The number of credits to be transferred following the completion of the practical training is…");

                InsertArticle(document, "Art 10. Report on the practical training",
                    "During the practical training, the training supervisor together with the supervising member of the teaching staff shall assess the trainee permanently, using an observation/assessment chart.");

                InsertArticle(document, "Art 11. Occupational health and safety",
                    "The practitioner attaches to this contract the proof of the medical insurance valid during the period and on the territory of the state where the practical training takes place.");

                InsertArticle(document, "Art 12. Discretionary conditions concerning the completion of the practical training",
                    "Allowances or bonuses awarded to the trainee: .......................");

                InsertArticle(document, "Art 13. Final provisions",
                    "Drawn up in three duplicates on: ...................");

                document.Save();

                Console.WriteLine("The document was generated successfully!");
            }
        }

        public void WriteConventionToWordRO()
        {
            using (var doc = DocX.Create(Stud.StudentName + Stud.StudentSurname + "-" + Comp.CompanyName + "ConventieRO.docx"))
            {
                doc.InsertParagraph("CONVENȚIE-CADRU")
                    .FontSize(14)
                    .Bold()
                    .Alignment = Alignment.center;

                doc.InsertParagraph("1. Universitatea Politehnica Timișoara")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Reprezentată de Rector, conf. univ. dr. ing. Florin DRĂGAN, "
                    + "cu sediul în TIMIȘOARA, Piața Victoriei, Nr. 2, cod 300006, telefon: 0256-403011, "
                    + "email: rector@upt.ro, cod unic de înregistrare: 4269282.")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("2. Partenerul de practică")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Reprezentat de: [" + Comp.CompanyOfficialPerson + "], "
                    + "cu sediul în [" + Comp.CompanyAddress + "], telefon: [" + Comp.CompanyPhoneNumber + "], e-mail: [" + Comp.CompanyEmail + "], cod fiscal: [" + Comp.CompanyFiscalCode + "], "
                    + "înregistrat la Registrul Comerțului cu numărul: [" + Comp.CompanyRegistrationNumber + "].")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("3. Studentul")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Nume și prenume: [" + Stud.GetFullName() + "], CNP: [" + Stud.StudentCNP + "], data nașterii: [" + Stud.StudentBornYear +"], "
                    + "cetățenie: [" + Stud.StudentCitizenship + "], adresă domiciliu: [" + Stud.StudentAddress + "].")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("Art. 1. Obiectul convenției-cadru")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Convenția-cadru stabilește modul în care se organizează și se desfășoară stagiul de practică "
                    + "pentru consolidarea cunoștințelor teoretice și formarea abilităților practice, aplicabile specializării studentului.")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("Art. 2. Statutul practicantului")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Practicantul rămâne student al Universității Politehnica Timișoara pe toată durata stagiului.")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("Art. 3. Durata și perioada desfășurării stagiului de practică")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Durata stagiului de practică este de [" + Information.HoursNumber + "], iar perioada de desfășurare este din [" + Information.StartDate + "] "
                    + "până la [" + Information.EndDate + "].")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("Art. 4. Plata și obligațiile sociale")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Stagiul de practică poate fi efectuat în cadrul unui contract de muncă sau nu, "
                    + "conform reglementărilor în vigoare.")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("Art. 5. Responsabilitățile practicantului")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Practicantul are obligația de a respecta programul de lucru și activitățile specificate în portofoliul de practică, "
                    + "inclusiv normele de sănătate și securitate.")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("Art. 6. Responsabilitățile partenerului de practică")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Partenerul de practică va desemna un tutore, care va superviza activitatea practică a studentului.")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("Art. 7. Obligațiile organizatorului de practică")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Organizatorul de practică va desemna un cadru didactic supervizor care va coordona stagiul de practică.")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("Art. 8. Persoane desemnate de organizatorul de practică și partenerul de practică")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Tutorele desemnat de partenerul de practică este [" + Tut.TuturName + "], funcția [" + Tut.TutorPosition + "], telefon [" + Tut.TutorPhoneNumber + "], e-mail [" + Tut.TutorEmail + "].")
                    .FontSize(12);

                doc.InsertParagraph("Cadrul didactic supervizor desemnat de organizatorul de practică este [" + Tss.TSSName + "], funcția [" + Tss.TSSPosition + "], "
                    + "telefon [" + Tss.TSSPhoneNumber + "], e-mail [" + Tss.TSSEmail + "].")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("Art. 9. Evaluarea stagiului de pregătire practică prin credite transferabile")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Numărul de credite transferabile care vor fi obținute este de [" + Information.CreditsNumber + "].")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("Art. 10. Raportul privind stagiul de pregătire practică")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Evaluarea va fi realizată pe baza raportului întocmit de tutore și al raportului final al studentului.")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("Art. 11. Sănătatea și securitatea în muncă")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Practicantul va prezenta dovada asigurării medicale valabile pentru perioada stagiului de practică.")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("Art. 12. Condiții facultative de desfășurare a stagiului de pregătire practică")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Practicul poate beneficia de indemnizație, gratificații, tichete de masă etc., în conformitate cu legislația în vigoare.")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("Art. 13. Prevederi finale")
                    .FontSize(12)
                    .Bold();

                doc.InsertParagraph("Convenția este întocmită în trei exemplare, semnătura părților fiind obligatorie.")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.InsertParagraph("Semnătura Universității Politehnica Timișoara: ___________________________")
                    .FontSize(12);

                doc.InsertParagraph("Semnătura Partenerului de practică: ________________________________")
                    .FontSize(12);

                doc.InsertParagraph("Semnătura Practicantului: _______________________________")
                    .FontSize(12);

                doc.InsertParagraph(Environment.NewLine);

                doc.Save();
                Console.WriteLine("Documentul a fost generat cu succes!");
            }
        }
    }

    public class ConventionsCollection
    {
        private List<PracticeConvention> Conventions;

        public List<PracticeConvention> GetConventions()
        {
            return Conventions;
        }

        public void SetConventions(List<PracticeConvention> conventions)
        {
            Conventions = conventions;
        }

        public ConventionsCollection(List<PracticeConvention> conventions)
        {
            Conventions = conventions;
        }

        public void AddConvention(PracticeConvention convention)
        {
            Conventions.Add(convention);
        }
        public void EditConvention(int id, PracticeConvention convention)
        {
            for (int i = 0; i < Conventions.Count; i++)
            {
                if (Conventions[i].Id == id)
                {
                    Conventions[i] = convention;
                    break;
                }
            }
        }
        public void ShowConventions()
        {
            Conventions.ForEach(convention => Console.WriteLine(convention.ToString()));
        }
        public void WriteConventionsToCSV(string filePath)
        {
            WriteToCSV(Conventions, filePath);
        }
        public void CreateWordDocumentForConvention(int id)
        {
            foreach (PracticeConvention convention in Conventions)
            {
                if (convention.Id == id)
                {
                    convention.WriteConventionToWordRO();
                    break;
                }
            }
        }
        public void createWordDocumentForEveryConvention()
        {
            foreach (PracticeConvention convention in Conventions)
            {
                convention.WriteConventionToWordRO();
            }
        }
    }

    static void Main(string[] args)
    {
        string filePath = @"C:\Users\Andrei\source\repos\ProjectCFLP\ProjectCFLP\dateProiect.csv";

        string WriteCsvPath = @"C:\Users\Andrei\source\repos\ProjectCFLP\ProjectCFLP\dateProiect1.csv";

        // Check if file exists
        if (!File.Exists(filePath))
        {
            Console.WriteLine("Fișierul nu există!");
            return;
        }

        try
        {
            // Read data from CSV
            List<Student> studs = ReadFromCSV<Student>(filePath).ToList();
            List<Company> comps = ReadFromCSV<Company>(filePath).ToList();
            List<Tutor> tuts =  ReadFromCSV<Tutor>(filePath).ToList();
            List<TeachingStaffSupervisor> tsss = ReadFromCSV<TeachingStaffSupervisor>(filePath).ToList();
            List<PracticeInformation> infos = ReadFromCSV<PracticeInformation>(filePath).ToList();

            if (studs.Count != comps.Count || studs.Count != tuts.Count || studs.Count != tsss.Count || studs.Count != infos.Count)
            {
                throw new Exception("The data lists have inconsistent lengths.");
            }

            // Create conventions collection
            ConventionsCollection cc = new ConventionsCollection(new List<PracticeConvention>());
            for (int i = 0; i < studs.Count(); i++)
            {
                PracticeConvention pc = new PracticeConvention(studs[i], comps[i], tuts[i], tsss[i], infos[i]);
                cc.AddConvention(pc);
            }

            cc.ShowConventions();

            // Write conventions to CSV
            cc.WriteConventionsToCSV(WriteCsvPath);

            // Generate Word documents for every convention
            cc.createWordDocumentForEveryConvention();

            PracticeConvention pc1 = cc.GetConventions()[0];
            pc1.WriteConventionToWordEN();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"A apărut o eroare: {ex.Message}");
        }
    }
    public static IEnumerable<T> ReadFromCSV<T>(string filePath)
    {
        using (var reader = new StreamReader(filePath))
        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
        {
            var allData = csv.GetRecords<T>().ToList();

            foreach (var data in allData)
            {
                foreach (var prop in typeof(T).GetProperties())
                {
                    var propName = prop.Name;
                    var propValue = prop.GetValue(data);
                    Console.WriteLine($"{propName}: {propValue}");
                }
                Console.WriteLine(new string('-', 40));
            }

            return allData;
        }
    }
    public static void WriteToCSV<T>(IEnumerable<T> data, string filePath)
    {
        using (var writer = new StreamWriter(filePath))
        using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
        {
            csv.WriteRecords(data);
        }
    }
    public static void InsertArticle(DocX document, string title, string content)
    {
        document.InsertParagraph(title)
               .FontSize(12)
               .Bold()
               .SpacingAfter(5);

        document.InsertParagraph(content)
               .FontSize(11)
               .SpacingAfter(15);
    }
}


// Embracing Possibility - A Woman's Complete Guide to Fertility, Nourishment & Wellness
// Professional e-book with cover, TOC, 4 parts, 12 chapters, backcover

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Docx;

public static class Program
{
    // ========================================================================
    // Color Scheme — Soft Berry Bloom (warm rose-plum, nurturing feminine)
    // ========================================================================
    private static class Colors
    {
        public const string Primary = "8e6b7a";       // Plum rose — headings
        public const string Secondary = "b8897a";      // Warm rose — accents
        public const string Accent = "c4a07a";         // Blush gold — highlights
        public const string Dark = "3d2f35";           // Deep plum brown — body text
        public const string Mid = "6b5259";            // Medium plum — secondary text
        public const string Light = "9e8a8f";          // Light mauve — captions
        public const string Border = "e0d5d0";         // Table borders
        public const string TableHeader = "f5efec";    // Table header bg
        public const string CalloutBg = "faf5f2";      // Callout box bg
    }

    private const int A4W = 11906;
    private const int A4H = 16838;
    private const long A4WE = 7560000L;
    private const long A4HE = 10692000L;

    public static void Main(string[] args)
    {
        string outputPath = args.Length > 0 ? args[0] : "/mnt/agents/output/Embracing_Possibility.docx";
        string bgDir = "/mnt/agents/output/bg";
        Generate(outputPath, bgDir);
    }

    public static void Generate(string outputPath, string bgDir)
    {
        using var doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        AddStyles(mainPart);
        AddNumbering(mainPart);

        var coverBgId = AddImage(mainPart, Path.Combine(bgDir, "cover_bg.png"));
        var backBgId = AddImage(mainPart, Path.Combine(bgDir, "backcover_bg.png"));

        uint prId = 1;
        AddCoverSection(body, coverBgId, ref prId);
        AddTocSection(body);
        AddContentSection(doc, body, mainPart, bgDir, ref prId);
        AddBackcoverSection(body, backBgId, ref prId);

        SetUpdateFieldsOnOpen(mainPart);
        doc.Save();
    }

    // ========================================================================
    // Styles
    // ========================================================================
    private static void AddStyles(MainDocumentPart mainPart)
    {
        var sp = mainPart.AddNewPart<StyleDefinitionsPart>();
        sp.Styles = new Styles();

        // Normal
        sp.Styles.Append(new Style(
            new StyleName { Val = "Normal" },
            new StyleParagraphProperties(
                new SpacingBetweenLines { After = "220", Line = "320", LineRule = LineSpacingRuleValues.Auto }),
            new StyleRunProperties(
                new RunFonts { Ascii = "Georgia", HighAnsi = "Georgia", EastAsia = "SimSun" },
                new FontSize { Val = "22" },
                new Color { Val = Colors.Dark })
        ) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });

        // Heading 1 — Part titles
        sp.Styles.Append(new Style(
            new StyleName { Val = "heading 1" }, new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepNext(), new KeepLines(), new PageBreakBefore(),
                new SpacingBetweenLines { Before = "400", After = "300" },
                new OutlineLevel { Val = 0 }),
            new StyleRunProperties(
                new Bold(), new FontSize { Val = "40" },
                new RunFonts { Ascii = "Georgia", HighAnsi = "Georgia", EastAsia = "SimSun" },
                new Color { Val = Colors.Primary },
                new Spacing { Val = 30 })
        ) { Type = StyleValues.Paragraph, StyleId = "Heading1" });

        // Heading 2 — Chapter titles
        sp.Styles.Append(new Style(
            new StyleName { Val = "heading 2" }, new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepNext(), new KeepLines(),
                new SpacingBetweenLines { Before = "400", After = "200" },
                new OutlineLevel { Val = 1 }),
            new StyleRunProperties(
                new Bold(), new FontSize { Val = "30" },
                new RunFonts { Ascii = "Georgia", HighAnsi = "Georgia", EastAsia = "SimSun" },
                new Color { Val = Colors.Secondary })
        ) { Type = StyleValues.Paragraph, StyleId = "Heading2" });

        // Heading 3 — Subsections
        sp.Styles.Append(new Style(
            new StyleName { Val = "heading 3" }, new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepNext(), new KeepLines(),
                new SpacingBetweenLines { Before = "280", After = "120" },
                new OutlineLevel { Val = 2 }),
            new StyleRunProperties(
                new Bold(), new FontSize { Val = "24" },
                new RunFonts { Ascii = "Georgia", HighAnsi = "Georgia", EastAsia = "SimSun" },
                new Color { Val = Colors.Mid })
        ) { Type = StyleValues.Paragraph, StyleId = "Heading3" });

        // Caption
        sp.Styles.Append(new Style(
            new StyleName { Val = "Caption" }, new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "60", After = "300" }),
            new StyleRunProperties(new Italic(), new Color { Val = Colors.Light }, new FontSize { Val = "20" })
        ) { Type = StyleValues.Paragraph, StyleId = "Caption" });

        // Quote
        sp.Styles.Append(new Style(
            new StyleName { Val = "Quote" }, new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new Indentation { Left = "720", Right = "720" },
                new SpacingBetweenLines { Before = "200", After = "200" }),
            new StyleRunProperties(new Italic(), new Color { Val = Colors.Mid }, new FontSize { Val = "22" })
        ) { Type = StyleValues.Paragraph, StyleId = "Quote" });

        // TOC styles
        sp.Styles.Append(CreateTocStyle("TOC1", "toc 1", true, "0", "240"));
        sp.Styles.Append(CreateTocStyle("TOC2", "toc 2", false, "360", "80"));
        sp.Styles.Append(CreateTocStyle("TOC3", "toc 3", false, "720", "60"));
    }

    private static Style CreateTocStyle(string id, string name, bool bold, string indent, string before)
    {
        var rpr = new StyleRunProperties(new Color { Val = bold ? Colors.Dark : Colors.Mid });
        if (bold) rpr.Append(new Bold());
        return new Style(
            new StyleName { Val = name }, new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new Tabs(new TabStop { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 9350 }),
                new SpacingBetweenLines { Before = before, After = "60" },
                new Indentation { Left = indent }),
            rpr
        ) { Type = StyleValues.Paragraph, StyleId = id };
    }

    // ========================================================================
    // Cover
    // ========================================================================
    private static void AddCoverSection(Body body, string coverBgId, ref uint prId)
    {
        body.Append(new Paragraph(new Run(CreateFloatingBackground(coverBgId, prId++, "CoverBg"))));
        body.Append(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "5500" }), new Run()));

        // Main title
        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "200" }),
            new Run(new RunProperties(
                    new FontSize { Val = "72" }, new Bold(),
                    new Color { Val = Colors.Dark },
                    new RunFonts { Ascii = "Georgia", HighAnsi = "Georgia" },
                    new Spacing { Val = 40 }),
                new Text("Embracing Possibility"))));

        // Subtitle
        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "400" }),
            new Run(new RunProperties(
                    new FontSize { Val = "28" },
                    new Color { Val = Colors.Mid },
                    new RunFonts { Ascii = "Georgia", HighAnsi = "Georgia" },
                    new Spacing { Val = 20 }),
                new Text("A Woman's Complete Guide to Fertility,"))));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "600" }),
            new Run(new RunProperties(
                    new FontSize { Val = "28" },
                    new Color { Val = Colors.Mid },
                    new RunFonts { Ascii = "Georgia", HighAnsi = "Georgia" },
                    new Spacing { Val = 20 }),
                new Text("Nourishment & Wellness Through Menopause and Beyond"))));

        // Decorative line
        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "600" }),
            new Run(new RunProperties(new Color { Val = Colors.Accent }, new FontSize { Val = "24" }),
                new Text("\u2014\u2014\u2014  \u2736  \u2014\u2014\u2014"))));

        // Tagline
        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "3000" }),
            new Run(new RunProperties(
                    new FontSize { Val = "22" }, new Italic(),
                    new Color { Val = Colors.Light }),
                new Text("Honest Science. Real Hope. Practical Wisdom."))));

        // Author line
        body.Append(new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
            new Run(new RunProperties(
                    new FontSize { Val = "20" },
                    new Color { Val = Colors.Light }),
                new Text("A comprehensive companion for every woman who still dares to dream"))));

        body.Append(new Paragraph(new ParagraphProperties(new SectionProperties(
            new TitlePage(),
            new SectionType { Val = SectionMarkValues.NextPage },
            new PageSize { Width = (UInt32Value)(uint)A4W, Height = (UInt32Value)(uint)A4H },
            new PageMargin { Top = 0, Right = 0, Bottom = 0, Left = 0, Header = 0, Footer = 0 }))));
    }

    // ========================================================================
    // TOC
    // ========================================================================
    private static void AddTocSection(Body body)
    {
        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "600", After = "400" }),
            new Run(new RunProperties(
                    new FontSize { Val = "36" }, new Bold(),
                    new Color { Val = Colors.Primary },
                    new Spacing { Val = 30 }),
                new Text("Contents"))));

        body.Append(new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines { After = "300" }),
            new Run(new RunProperties(new Color { Val = Colors.Light }, new FontSize { Val = "18" }, new Italic()),
                new Text("Right-click and select \"Update Field\" to refresh page numbers after editing"))));

        body.Append(new Paragraph(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" TOC \\o \"1-3\" \\h \\z \\u ") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate })));

        string[,] toc = {
            { "Introduction: A Letter to Every Woman Who Still Dreams", "1", "3" },
            { "Part One: Understanding Your Changing Fertility", "1", "5" },
            { "Chapter 1: The Truth About Perimenopause and Your Fertility", "2", "5" },
            { "Chapter 2: Reading Your Body's Signals", "2", "8" },
            { "Chapter 3: When to Seek Help", "2", "12" },
            { "Part Two: The Path of IVF and Donor Eggs", "1", "16" },
            { "Chapter 4: Understanding IVF After Menopause", "2", "16" },
            { "Chapter 5: The Gift of Donor Eggs", "2", "20" },
            { "Chapter 6: Preparing Your Body for the Journey", "2", "24" },
            { "Part Three: Nourishing Your Body, Nurturing Your Dreams", "1", "28" },
            { "Chapter 7: Foods That Support Hormonal Harmony", "2", "28" },
            { "Chapter 8: The Fertility-Boosting Plate", "2", "32" },
            { "Chapter 9: Timing, Movement, and Rest", "2", "36" },
            { "Part Four: Emotional Wellness and Your Support Circle", "1", "40" },
            { "Chapter 10: Honoring Your Emotional Journey", "2", "40" },
            { "Chapter 11: Building Your Support System", "2", "44" },
            { "Chapter 12: Stories of Hope and Possibility", "2", "48" },
            { "Closing Words: Your Journey Forward", "1", "52" },
        };
        for (int i = 0; i < toc.GetLength(0); i++)
            body.Append(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = $"TOC{toc[i, 1]}" }),
                new Run(new Text(toc[i, 0])), new Run(new TabChar()), new Run(new Text(toc[i, 2]))));

        body.Append(new Paragraph(new Run(new FieldChar { FieldCharType = FieldCharValues.End })));

        body.Append(new Paragraph(new ParagraphProperties(new SectionProperties(
            new SectionType { Val = SectionMarkValues.NextPage },
            new PageSize { Width = (UInt32Value)(uint)A4W, Height = (UInt32Value)(uint)A4H },
            new PageMargin { Top = 1800, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720 }))));
    }

    // ========================================================================
    // Content
    // ========================================================================
    private static void AddContentSection(WordprocessingDocument doc, Body body,
        MainDocumentPart mainPart, string bgDir, ref uint prId)
    {
        // Header
        var headerPart = mainPart.AddNewPart<HeaderPart>();
        var headerId = mainPart.GetIdOfPart(headerPart);
        var bodyBgPath = Path.Combine(bgDir, "body_bg.png");
        if (File.Exists(bodyBgPath))
        {
            var headerImagePart = headerPart.AddImagePart(ImagePartType.Png);
            using (var stream = new FileStream(bodyBgPath, FileMode.Open))
                headerImagePart.FeedData(stream);
            var headerImageId = headerPart.GetIdOfPart(headerImagePart);
            headerPart.Header = new Header(
                new Paragraph(new Run(CreateFloatingBackground(headerImageId, prId++, "BodyBg"))),
                new Paragraph(
                    new ParagraphProperties(new Justification { Val = JustificationValues.Right }),
                    new Run(new RunProperties(new FontSize { Val = "18" }, new Italic(), new Color { Val = Colors.Light }),
                        new Text("Embracing Possibility"))));
        }
        else
        {
            headerPart.Header = new Header(new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Right }),
                new Run(new RunProperties(new FontSize { Val = "18" }, new Italic(), new Color { Val = Colors.Light }),
                    new Text("Embracing Possibility"))));
        }

        // Footer
        var footerPart = mainPart.AddNewPart<FooterPart>();
        var footerId = mainPart.GetIdOfPart(footerPart);
        var fp = new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }));
        fp.Append(new Run(new RunProperties(new FontSize { Val = "18" }, new Color { Val = Colors.Light }),
            new FieldChar { FieldCharType = FieldCharValues.Begin }));
        fp.Append(new Run(new RunProperties(new FontSize { Val = "18" }, new Color { Val = Colors.Light }),
            new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve }));
        fp.Append(new Run(new RunProperties(new FontSize { Val = "18" }, new Color { Val = Colors.Light }),
            new FieldChar { FieldCharType = FieldCharValues.Separate }));
        fp.Append(new Run(new RunProperties(new FontSize { Val = "18" }, new Color { Val = Colors.Light }),
            new Text("1")));
        fp.Append(new Run(new RunProperties(new FontSize { Val = "18" }, new Color { Val = Colors.Light }),
            new FieldChar { FieldCharType = FieldCharValues.End }));
        footerPart.Footer = new Footer(fp);

        // ==================== INTRODUCTION ====================
        body.Append(CreateHeading1("Introduction", "_TocIntro"));
        body.Append(CreateHeading2("A Letter to Every Woman Who Still Dreams"));

        body.Append(CreateParagraph("If you are holding this book, chances are you carry a dream that society has told you to release. Perhaps you are in your late thirties, your forties, or beyond, and the word \"menopause\" has entered your vocabulary with a weight you never expected. Perhaps you have heard whispers that your window has closed, that your body has moved on to a new chapter where motherhood is no longer possible."));

        body.Append(CreateQuote("\"The question isn't whether the path is easy. The question is whether the dream is worth the journey.\""));

        body.Append(CreateParagraph("This book exists because I believe in the power of informed hope. Not blind optimism, not false promises, but the kind of hope that is grounded in science, nourished by wisdom, and carried forward by women who refuse to let their age define their possibilities."));

        body.Append(CreateParagraph("Inside these pages, you will find honest truths about fertility during perimenopause and after menopause. You will learn about the remarkable advances in reproductive medicine that have made motherhood possible for women well into their fifties and beyond. You will discover the foods, habits, and mindset shifts that can support your body through this journey. And perhaps most importantly, you will find companionship for the emotional road you walk."));

        body.Append(CreateHeading3("What This Book Is — And What It Is Not"));

        body.Append(CreateParagraph("This is not a book of miracle cures or guaranteed formulas. Fertility is deeply individual, influenced by genetics, health history, and circumstances unique to each woman. What this book offers is a comprehensive, compassionate guide to understanding your options, optimizing your health, and making informed decisions with clarity and confidence."));

        body.Append(CreateBulletParagraph("\u2022 Honest, science-based information about fertility during and after menopause"));
        body.Append(CreateBulletParagraph("\u2022 Clear explanations of IVF, donor eggs, and reproductive technologies"));
        body.Append(CreateBulletParagraph("\u2022 Practical nutrition and lifestyle guidance backed by research"));
        body.Append(CreateBulletParagraph("\u2022 Emotional support and strategies for navigating this journey"));
        body.Append(CreateBulletParagraph("\u2022 Realistic timelines, financial considerations, and decision-making frameworks"));

        body.Append(CreateHeading3("How to Use This Book"));

        body.Append(CreateParagraph("You may read this book from cover to cover, or you may turn directly to the sections that speak to where you are right now. Part One will help you understand your current fertility status. Part Two explores assisted reproductive technologies. Part Three offers practical nutrition and lifestyle guidance. Part Four addresses the emotional and spiritual dimensions of this journey."));

        body.Append(CreateParagraph("Wherever you begin, know this: you are not alone. Thousands of women have walked this path before you, and thousands more will follow. Your desire for motherhood is not something to be ashamed of or dismissed. It is a testament to the life force within you, and it deserves to be honored with honest information and compassionate guidance."));

        body.Append(CreateCallout("Let this book be your companion, your reference, and your source of strength as you explore what is possible. Your journey is uniquely yours, but you do not have to walk it alone."));

        // ==================== PART ONE ====================
        body.Append(CreateHeading1("Part One: Understanding Your Changing Fertility", "_TocPart1"));

        body.Append(CreatePartIntro("Before we explore solutions, we must understand the landscape. This section offers a clear-eyed look at what happens to fertility during perimenopause, how to read your body's signals, and when to seek professional guidance."));

        // Chapter 1
        body.Append(CreateHeading2("Chapter 1: The Truth About Perimenopause and Your Fertility"));

        body.Append(CreateChapterOpener("There is a season in every woman's life when the rhythm of her body begins to shift. Understanding this transition is the first step toward making empowered choices about your fertility."));

        body.Append(CreateHeading3("What Is Perimenopause, Really?"));

        body.Append(CreateParagraph("Perimenopause is the transitional period that leads up to menopause, typically beginning in a woman's late thirties to mid-forties, though for some it starts earlier. During this time, your ovaries gradually begin to produce less estrogen, and your menstrual cycles may become irregular. This phase can last anywhere from a few months to over a decade."));

        body.Append(CreateParagraph("The critical thing to understand is that perimenopause is not menopause. You are still ovulating, still producing eggs, and yes — still capable of conceiving. The window may be narrowing, but for many women, it remains open longer than they have been led to believe."));

        // Fertility timeline table
        body.Append(CreateHeading3("Your Fertility Timeline: What the Research Shows"));
        body.Append(CreateFertilityTable());
        body.Append(CreateCaption("Table 1: Average Fertility Rates by Age Group"));

        body.Append(CreateHeading3("The Biological Reality"));

        body.Append(CreateParagraph("Women are born with approximately one to two million eggs. By puberty, this number has dropped to about 300,000. Each menstrual cycle, roughly one thousand eggs begin maturing, but only one (or occasionally two) will fully develop and be released. The rest are reabsorbed by the body."));

        body.Append(CreateParagraph("By age 35, the decline in both egg quantity and quality becomes more pronounced. By 40, the average woman has a 5% chance of conceiving per cycle. But here is what is often left unsaid: these are averages. Individual fertility varies enormously based on genetics, overall health, ovarian reserve, and lifestyle factors."));

        body.Append(CreateHeading3("Signs Your Fertility Is Transitioning"));

        body.Append(CreateBulletParagraph("\u2022 Changes in cycle length — cycles may become shorter or longer than your usual pattern"));
        body.Append(CreateBulletParagraph("\u2022 Heavier or lighter menstrual flow than you typically experience"));
        body.Append(CreateBulletParagraph("\u2022 Hot flashes or night sweats, even before your period stops completely"));
        body.Append(CreateBulletParagraph("\u2022 Mood changes, irritability, or anxiety that feels different from your normal patterns"));
        body.Append(CreateBulletParagraph("\u2022 Changes in cervical mucus consistency and quantity"));
        body.Append(CreateBulletParagraph("\u2022 Sleep disturbances, particularly waking in the early morning hours"));
        body.Append(CreateBulletParagraph("\u2022 Vaginal dryness or discomfort during intimacy"));

        body.Append(CreateHeading3("Testing Your Ovarian Reserve"));

        body.Append(CreateParagraph("If you are considering pregnancy during perimenopause, the first step is comprehensive fertility testing. Your doctor can order several key tests to assess your ovarian reserve — the quantity and quality of eggs remaining in your ovaries."));

        body.Append(CreateHeading3("Key Fertility Tests"));
        body.Append(CreateTestsTable());
        body.Append(CreateCaption("Table 2: Essential Fertility Tests for Women Over 40"));

        body.Append(CreateCallout("Important: These tests provide a snapshot, not a verdict. Many women with lower-than-average ovarian reserve have successfully conceived, while others with seemingly good numbers may still face challenges. Use these tests as guidance, not prophecy."));

        // Chapter 2
        body.Append(CreateHeading2("Chapter 2: Reading Your Body's Signals"));

        body.Append(CreateChapterOpener("Your body speaks to you in subtle whispers and clear signs. Learning to read these signals can help you identify your most fertile windows and understand what your body needs."));

        body.Append(CreateHeading3("Understanding Ovulation After 40"));

        body.Append(CreateParagraph("Contrary to popular belief, many women continue to ovulate regularly throughout their early to mid-forties. The key difference is that ovulation may not occur every single cycle, and the timing may become less predictable. This is why learning to read your body's fertility signals becomes even more crucial as you age."));

        body.Append(CreateHeading3("Tracking Your Basal Body Temperature"));

        body.Append(CreateParagraph("Your basal body temperature (BBT) is your body's resting temperature, taken immediately upon waking before any activity. During the first half of your cycle, estrogen keeps your temperature slightly lower. After ovulation, progesterone causes a measurable rise of 0.3 to 0.5 degrees Fahrenheit."));

        body.Append(CreateBulletParagraph("\u2022 Use a digital thermometer designed for BBT tracking — it will measure to the hundredth of a degree"));
        body.Append(CreateBulletParagraph("\u2022 Take your temperature at the same time every morning, before getting out of bed"));
        body.Append(CreateBulletParagraph("\u2022 Record your readings on a chart or fertility tracking app"));
        body.Append(CreateBulletParagraph("\u2022 Look for a sustained temperature rise of at least three days to confirm ovulation"));
        body.Append(CreateBulletParagraph("\u2022 Keep in mind that illness, poor sleep, and alcohol can affect your readings"));

        body.Append(CreateHeading3("Cervical Mucus: Your Body's Fertility Sign"));

        body.Append(CreateParagraph("Cervical mucus is one of the most reliable indicators of approaching ovulation. As estrogen rises before ovulation, your cervical mucus changes from thick and sticky to thin, slippery, and stretchy — similar to raw egg whites. This fertile-quality mucus helps sperm survive and travel toward the egg."));

        body.Append(CreateQuote("\"Your cervical mucus is like a fertility highway. When it becomes clear and stretchy, your body is literally preparing a path for sperm to reach your egg.\""));

        body.Append(CreateHeading3("Ovulation Predictor Kits"));

        body.Append(CreateParagraph("Ovulation predictor kits (OPKs) detect the surge of luteinizing hormone (LH) that triggers ovulation, typically 24 to 36 hours before the egg is released. For women over 40, these kits can be particularly helpful because ovulation timing becomes less predictable."));

        body.Append(CreateBulletParagraph("\u2022 Begin testing a few days before you expect to ovulate"));
        body.Append(CreateBulletParagraph("\u2022 Test at the same time each day, preferably in the afternoon"));
        body.Append(CreateBulletParagraph("\u2022 A positive result means ovulation will likely occur within the next day and a half"));
        body.Append(CreateBulletParagraph("\u2022 Some women experience multiple LH surges before a true ovulation — don't be discouraged"));

        body.Append(CreateHeading3("The Fertility Awareness Method"));

        body.Append(CreateParagraph("Fertility awareness involves tracking multiple signs together to identify your fertile window. When BBT, cervical mucus, and OPK results are combined, they create a comprehensive picture of your cycle. Consider working with a certified fertility awareness educator for personalized guidance."));

        // Chapter 3
        body.Append(CreateHeading2("Chapter 3: When to Seek Help"));

        body.Append(CreateChapterOpener("Knowing when to transition from natural conception attempts to professional assistance can save precious time and emotional energy. This chapter will help you make that decision with clarity."));

        body.Append(CreateHeading3("The Six-Month Rule for Women Over 40"));

        body.Append(CreateParagraph("For women under 35, doctors typically recommend trying to conceive for a full year before seeking fertility evaluation. However, for women over 40, this timeline shortens dramatically to just six months. After 45, most fertility specialists recommend immediate evaluation if you are trying to conceive."));

        body.Append(CreateParagraph("This compressed timeline exists for a reason. Every month matters when your ovarian reserve is declining. Early intervention can mean the difference between using your own eggs and needing donor eggs, or between a simple treatment protocol and a more complex one."));

        body.Append(CreateHeading3("Choosing a Fertility Specialist"));

        body.Append(CreateParagraph("Not all fertility clinics are created equal, and finding the right doctor can significantly impact your experience and outcomes. Here is what to look for:"));

        body.Append(CreateBulletParagraph("\u2022 Board certification in reproductive endocrinology and infertility (REI)"));
        body.Append(CreateBulletParagraph("\u2022 Experience specifically with women in their forties and beyond"));
        body.Append(CreateBulletParagraph("\u2022 A clinic with published success rates for your age group"));
        body.Append(CreateBulletParagraph("\u2022 Availability of donor egg programs if that may be part of your journey"));
        body.Append(CreateBulletParagraph("\u2022 A communication style that makes you feel heard and respected"));

        body.Append(CreateHeading3("Questions to Ask at Your First Appointment"));
        body.Append(CreateQuestionsTable());
        body.Append(CreateCaption("Table 3: Essential Questions for Your Fertility Consultation"));

        body.Append(CreateHeading3("Understanding Your Treatment Options"));

        body.Append(CreateParagraph("Depending on your test results and overall health, your doctor may recommend several pathways. Timed intercourse with ovulation monitoring is the least invasive option. Oral medications like letrozole or clomiphene can stimulate egg development. Injectable fertility drugs offer stronger stimulation. Intrauterine insemination (IUI) places prepared sperm directly into the uterus. In vitro fertilization (IVF) offers the highest success rates, particularly for women over 40."));

        body.Append(CreateCallout("Remember: Your doctor is your partner in this journey, not the director of it. Ask questions, seek second opinions when needed, and trust your intuition about what feels right for your body and your life."));

        // ==================== PART TWO ====================
        body.Append(CreateHeading1("Part Two: The Path of IVF and Donor Eggs", "_TocPart2"));

        body.Append(CreatePartIntro("When natural conception is no longer possible, modern reproductive medicine offers remarkable pathways to motherhood. This section explores IVF after menopause, the transformative option of donor eggs, and how to prepare your body for these journeys."));

        // Chapter 4
        body.Append(CreateHeading2("Chapter 4: Understanding IVF After Menopause"));

        body.Append(CreateChapterOpener("In vitro fertilization has rewritten the rules of what is possible. For women navigating menopause, understanding this technology opens doors that were firmly shut just one generation ago."));

        body.Append(CreateHeading3("How IVF Works: A Clear Explanation"));

        body.Append(CreateParagraph("In vitro fertilization literally means \"fertilization in glass.\" The process involves stimulating the ovaries to produce multiple eggs, retrieving those eggs, fertilizing them with sperm in a laboratory, and then transferring the resulting embryo into the uterus. For women in perimenopause or postmenopause, the process adapts to work with your body's current capabilities."));

        body.Append(CreateHeading3("The IVF Process Step by Step"));

        body.Append(CreateBulletParagraph("Step 1: Ovarian Stimulation — Fertility medications are used to encourage your ovaries to produce multiple mature eggs. For women over 40, protocols are carefully tailored to maximize quality over quantity."));
        body.Append(CreateBulletParagraph("Step 2: Monitoring — Through blood tests and ultrasounds, your doctor tracks follicle development to determine the optimal time for egg retrieval."));
        body.Append(CreateBulletParagraph("Step 3: Egg Retrieval — A minor surgical procedure uses ultrasound guidance to collect mature eggs from the ovaries. This takes about 20 minutes and is performed under sedation."));
        body.Append(CreateBulletParagraph("Step 4: Fertilization — Collected eggs are combined with sperm in the laboratory. Intracytoplasmic sperm injection (ICSI) may be used, where a single sperm is directly injected into each egg."));
        body.Append(CreateBulletParagraph("Step 5: Embryo Development — Fertilized eggs are monitored for 5-6 days as they develop into blastocysts. The strongest embryos are selected for transfer."));
        body.Append(CreateBulletParagraph("Step 6: Embryo Transfer — One or more embryos are placed into the uterus through a thin catheter. This simple procedure requires no anesthesia."));
        body.Append(CreateBulletParagraph("Step 7: The Two-Week Wait — Approximately two weeks after transfer, a blood test determines whether pregnancy has been achieved."));

        body.Append(CreateHeading3("Realistic Success Rates"));
        body.Append(CreateIVFSuccessTable());
        body.Append(CreateCaption("Table 4: IVF Success Rates Using Own Eggs (National Averages)"));

        body.Append(CreateHeading3("Why Age Matters So Much"));

        body.Append(CreateParagraph("The relationship between age and IVF success is primarily about egg quality. As women age, chromosomal abnormalities in eggs become more common. A 30-year-old's eggs have roughly a 20-25% rate of chromosomal abnormalities. By age 45, that rate climbs to 75-90%. This is why miscarriage rates are higher for older women and why genetic testing of embryos becomes increasingly important."));

        body.Append(CreateHeading3("Embryo Genetic Testing"));

        body.Append(CreateParagraph("Preimplantation genetic testing (PGT) allows embryos to be screened for chromosomal abnormalities before transfer. For women over 40, this testing is particularly valuable because it helps identify the embryos most likely to result in a healthy, full-term pregnancy. While it adds cost and time to the process, it can significantly improve success rates per transfer and reduce the risk of miscarriage."));

        // Chapter 5
        body.Append(CreateHeading2("Chapter 5: The Gift of Donor Eggs"));

        body.Append(CreateChapterOpener("For many women, donor eggs represent not a second choice, but a beautiful pathway to the family they have always envisioned. Understanding this option can transform despair into renewed hope."));

        body.Append(CreateHeading3("Why Donor Eggs Change Everything"));

        body.Append(CreateParagraph("Using donor eggs dramatically increases IVF success rates for women over 42 and makes pregnancy possible even after natural menopause. With donor eggs, the success rate per embryo transfer can exceed 50%, regardless of the recipient's age. This is because the egg's age — not the uterus's age — determines the vast majority of pregnancy outcomes."));

        body.Append(CreateQuote("\"A donor egg pregnancy is still your pregnancy. Every heartbeat, every kick, every moment of connection happens in your body. The nurturing environment you provide shapes your baby's development in profound ways that science is only beginning to understand.\""));

        body.Append(CreateHeading3("Choosing an Egg Donor"));

        body.Append(CreateParagraph("Egg donors can be known (friends or family members) or anonymous (selected through an agency or clinic database). Most programs offer detailed profiles including physical characteristics, education, medical history, and sometimes personal essays or photos."));

        body.Append(CreateHeading3("The Donor Egg IVF Process"));

        body.Append(CreateBulletParagraph("\u2022 The egg donor undergoes ovarian stimulation and egg retrieval"));
        body.Append(CreateBulletParagraph("\u2022 Retrieved eggs are fertilized with your partner's sperm (or donor sperm)"));
        body.Append(CreateBulletParagraph("\u2022 Meanwhile, you take medications to prepare your uterine lining"));
        body.Append(CreateBulletParagraph("\u2022 When embryos reach the blastocyst stage, one is selected for transfer"));
        body.Append(CreateBulletParagraph("\u2022 The embryo is transferred to your uterus at the optimal time"));
        body.Append(CreateBulletParagraph("\u2022 Any remaining viable embryos can be frozen for future use"));

        body.Append(CreateHeading3("Legal and Ethical Considerations"));

        body.Append(CreateParagraph("Laws regarding egg donation vary by country and state. In most jurisdictions, egg donors have no parental rights or responsibilities when working through a licensed clinic. Legal agreements should clearly establish parentage, confidentiality, and any future contact arrangements. Consulting with a reproductive attorney is strongly recommended."));

        body.Append(CreateHeading3("Financial Investment"));
        body.Append(CreateCostTable());
        body.Append(CreateCaption("Table 5: Estimated Costs for Donor Egg IVF (US Averages)"));

        // Chapter 6
        body.Append(CreateHeading2("Chapter 6: Preparing Your Body for the Journey"));

        body.Append(CreateChapterOpener("Whether pursuing IVF with your own eggs or preparing for donor egg transfer, the condition of your uterus and overall health significantly impacts success. This chapter covers the essential preparation steps."));

        body.Append(CreateHeading3("The Uterus Does Not Age Like the Ovaries"));

        body.Append(CreateParagraph("Here is the encouraging truth: while egg quality declines significantly with age, the uterus remains remarkably receptive well into a woman's fifties and even sixties. Research has shown that women in their fifties can carry pregnancies to term with success rates similar to younger women when using donor eggs. The key is ensuring your uterine environment is healthy and receptive."));

        body.Append(CreateHeading3("Optimizing Uterine Health"));

        body.Append(CreateBulletParagraph("\u2022 Saline sonogram or hysteroscopy to evaluate the uterine cavity"));
        body.Append(CreateBulletParagraph("\u2022 Treatment of any fibroids, polyps, or scar tissue that could affect implantation"));
        body.Append(CreateBulletParagraph("\u2022 Management of endometriosis if present"));
        body.Append(CreateBulletParagraph("\u2022 Treatment of any chronic conditions like thyroid disorders or diabetes"));
        body.Append(CreateBulletParagraph("\u2022 Evaluation and treatment of autoimmune factors if indicated"));

        body.Append(CreateHeading3("The Three-Month Preparation Window"));

        body.Append(CreateParagraph("Most fertility specialists recommend at least three months of focused preparation before an IVF cycle. This timing aligns with the approximately 90-day development window of eggs and the cycle of uterine lining regeneration. Use this time to optimize your nutrition, establish an exercise routine, manage stress, and ensure all medical conditions are well-controlled."));

        body.Append(CreateHeading3("Medical Clearances for Older Mothers-to-Be"));

        body.Append(CreateParagraph("If you are over 45 and pursuing pregnancy, your doctor will likely request additional medical evaluations to ensure you can safely carry a pregnancy. These may include a cardiac stress test, diabetes screening, blood pressure monitoring, and assessment of any existing medical conditions."));

        body.Append(CreateCallout("Pregnancy after 45 is considered high-risk, but that does not mean it is impossible or inadvisable. With proper medical care and monitoring, many women in their late forties and fifties have healthy pregnancies and deliveries. The key is being informed, prepared, and supported by an experienced medical team."));

        // ==================== PART THREE ====================
        body.Append(CreateHeading1("Part Three: Nourishing Your Body, Nurturing Your Dreams", "_TocPart3"));

        body.Append(CreatePartIntro("What you eat, how you move, and the quality of your rest all influence your hormonal balance, egg quality, and uterine receptivity. This section offers practical, evidence-based guidance for supporting your fertility through nutrition and lifestyle."));

        // Chapter 7
        body.Append(CreateHeading2("Chapter 7: Foods That Support Hormonal Harmony"));

        body.Append(CreateChapterOpener("Food is information for your body. Every bite you take sends signals to your hormones, your ovaries, and your uterus. Learning to choose foods that support fertility can be both empowering and delicious."));

        body.Append(CreateHeading3("The Anti-Inflammatory Approach"));

        body.Append(CreateParagraph("Chronic inflammation has been linked to decreased fertility, poor egg quality, and implantation failure. An anti-inflammatory diet emphasizes whole, unprocessed foods rich in antioxidants and omega-3 fatty acids while minimizing inflammatory triggers like refined sugars, processed meats, and trans fats."));

        body.Append(CreateHeading3("Building Your Fertility Plate"));
        body.Append(CreateNutritionTable());
        body.Append(CreateCaption("Table 6: The Fertility-Boosting Food Framework"));

        body.Append(CreateHeading3("Key Nutrients for Egg Quality"));

        body.Append(CreateBulletParagraph("\u2022 Coenzyme Q10 (CoQ10) — 200-600mg daily. This powerful antioxidant supports mitochondrial function in eggs, which is crucial for energy production during fertilization and early embryo development."));
        body.Append(CreateBulletParagraph("\u2022 DHEA — 25mg three times daily (under doctor supervision). Some studies suggest DHEA supplementation may improve ovarian response in women with diminished reserve."));
        body.Append(CreateBulletParagraph("\u2022 Omega-3 Fatty Acids — 1-2g daily from fish oil or algae sources. Essential for hormone production and reducing inflammation."));
        body.Append(CreateBulletParagraph("\u2022 Vitamin D — 2000-4000 IU daily. Adequate vitamin D levels are associated with better IVF outcomes and overall fertility."));
        body.Append(CreateBulletParagraph("\u2022 Folate — 800mcg daily from methylfolate (the active form). Critical for DNA synthesis and early fetal development."));
        body.Append(CreateBulletParagraph("\u2022 Antioxidants — Vitamin C, Vitamin E, and alpha-lipoic acid protect eggs from oxidative stress."));

        body.Append(CreateHeading3("Foods to Embrace"));

        body.Append(CreateBulletParagraph("\u2022 Leafy greens (spinach, kale, Swiss chard) — folate, iron, calcium"));
        body.Append(CreateBulletParagraph("\u2022 Fatty fish (salmon, sardines, mackerel) — omega-3s, vitamin D, protein"));
        body.Append(CreateBulletParagraph("\u2022 Colorful vegetables (beets, bell peppers, sweet potatoes) — antioxidants"));
        body.Append(CreateBulletParagraph("\u2022 Berries (blueberries, raspberries, strawberries) — anthocyanins, vitamin C"));
        body.Append(CreateBulletParagraph("\u2022 Nuts and seeds (walnuts, flaxseeds, pumpkin seeds) — healthy fats, zinc"));
        body.Append(CreateBulletParagraph("\u2022 Legumes (lentils, chickpeas, black beans) — plant protein, fiber, iron"));
        body.Append(CreateBulletParagraph("\u2022 Whole grains (quinoa, oats, brown rice) — B vitamins, fiber"));
        body.Append(CreateBulletParagraph("\u2022 Fermented foods (yogurt, kefir, sauerkraut) — gut health, immune support"));

        body.Append(CreateHeading3("Foods to Limit or Avoid"));

        body.Append(CreateBulletParagraph("\u2022 Trans fats and partially hydrogenated oils"));
        body.Append(CreateBulletParagraph("\u2022 Highly processed and fast foods"));
        body.Append(CreateBulletParagraph("\u2022 Excessive caffeine (limit to 200mg daily — about one 12-ounce coffee)"));
        body.Append(CreateBulletParagraph("\u2022 Alcohol (best avoided entirely during preconception and treatment)"));
        body.Append(CreateBulletParagraph("\u2022 High-mercury fish (shark, swordfish, king mackerel, tilefish)"));
        body.Append(CreateBulletParagraph("\u2022 Non-organic produce with high pesticide residues (the \"Dirty Dozen\")"));
        body.Append(CreateBulletParagraph("\u2022 Artificial sweeteners and high-fructose corn syrup"));

        // Chapter 8
        body.Append(CreateHeading2("Chapter 8: The Fertility-Boosting Plate"));

        body.Append(CreateChapterOpener("Knowing what to eat is one thing. Putting it together into meals you actually enjoy is another. This chapter brings the principles to life with practical meal planning and delicious combinations."));

        body.Append(CreateHeading3("The Ideal Fertility Plate"));

        body.Append(CreateParagraph("Aim for each meal to contain: half your plate in colorful vegetables, a quarter in high-quality protein, a quarter in whole grains or starchy vegetables, and a generous source of healthy fats. This balance stabilizes blood sugar, supports hormone production, and provides the nutrients your reproductive system needs."));

        body.Append(CreateSampleMealsTable());
        body.Append(CreateCaption("Table 7: Sample Fertility-Supporting Meals"));

        body.Append(CreateHeading3("The Importance of Meal Timing"));

        body.Append(CreateParagraph("When you eat matters almost as much as what you eat. Aim to eat breakfast within an hour of waking to stabilize blood sugar and support cortisol rhythm. Space meals evenly throughout the day — typically every 3-4 hours — to maintain steady energy and hormone levels. Avoid eating large meals close to bedtime, as this can disrupt sleep quality."));

        body.Append(CreateHeading3("Hydration for Fertility"));

        body.Append(CreateParagraph("Proper hydration supports cervical mucus production, hormone transport, and cellular function throughout your reproductive system. Aim for 8-10 glasses of filtered water daily. Herbal teas like red raspberry leaf, nettle, and peppermint can contribute to your fluid intake while offering additional nutritional benefits."));

        body.Append(CreateHeading3("Supplement Timing"));

        body.Append(CreateParagraph("For optimal absorption, take fat-soluble vitamins (D, E, CoQ10) with meals containing healthy fats. Take iron supplements with vitamin C-rich foods and away from calcium, which can inhibit absorption. B vitamins are best taken in the morning as they can be energizing. Always discuss supplement timing with your healthcare provider."));

        // Chapter 9
        body.Append(CreateHeading2("Chapter 9: Timing, Movement, and Rest"));

        body.Append(CreateChapterOpener("Your body was designed to move, to rest, and to follow natural rhythms. In our modern world, these fundamental needs are often neglected. Reclaiming them is one of the most powerful fertility supports available."));

        body.Append(CreateHeading3("Exercise: Finding the Sweet Spot"));

        body.Append(CreateParagraph("Moderate exercise supports fertility by reducing stress, improving insulin sensitivity, and promoting healthy hormone balance. However, excessive high-intensity exercise can actually suppress reproductive function. The key is finding your personal sweet spot."));

        body.Append(CreateBulletParagraph("\u2022 Aim for 30 minutes of moderate activity most days of the week"));
        body.Append(CreateBulletParagraph("\u2022 Walking, swimming, yoga, and Pilates are excellent fertility-friendly options"));
        body.Append(CreateBulletParagraph("\u2022 Strength training 2-3 times weekly supports hormone balance and bone density"));
        body.Append(CreateBulletParagraph("\u2022 Avoid excessive high-intensity workouts during fertility treatment cycles"));
        body.Append(CreateBulletParagraph("\u2022 Listen to your body — rest days are as important as active days"));

        body.Append(CreateHeading3("The Power of Sleep"));

        body.Append(CreateParagraph("Sleep is when your body repairs, regenerates, and rebalances hormones. Growth hormone, crucial for egg quality, is primarily released during deep sleep. Melatonin, produced during darkness, acts as a powerful antioxidant that protects eggs from oxidative damage."));

        body.Append(CreateBulletParagraph("\u2022 Aim for 7-9 hours of sleep per night"));
        body.Append(CreateBulletParagraph("\u2022 Maintain a consistent sleep schedule, even on weekends"));
        body.Append(CreateBulletParagraph("\u2022 Create a dark, cool sleeping environment (65-68 degrees Fahrenheit ideal)"));
        body.Append(CreateBulletParagraph("\u2022 Avoid screens for at least one hour before bed — blue light suppresses melatonin"));
        body.Append(CreateBulletParagraph("\u2022 Consider a magnesium supplement before bed to support relaxation"));
        body.Append(CreateBulletParagraph("\u2022 Establish a calming bedtime routine to signal your body it's time to rest"));

        body.Append(CreateHeading3("Stress Management: The Non-Negotiable"));

        body.Append(CreateParagraph("Chronic stress elevates cortisol, which can disrupt ovulation, impair implantation, and reduce IVF success rates. While eliminating all stress is impossible, developing effective management strategies is essential for your fertility journey."));

        body.Append(CreateBulletParagraph("\u2022 Mindfulness meditation — even 10 minutes daily can measurably reduce cortisol"));
        body.Append(CreateBulletParagraph("\u2022 Yoga and tai chi combine movement with breath awareness and stress reduction"));
        body.Append(CreateBulletParagraph("\u2022 Acupuncture has shown promising results for improving IVF outcomes"));
        body.Append(CreateBulletParagraph("\u2022 Journaling can help process emotions and reduce mental rumination"));
        body.Append(CreateBulletParagraph("\u2022 Spending time in nature lowers blood pressure and stress hormones"));
        body.Append(CreateBulletParagraph("\u2022 Connecting with supportive friends and community provides emotional regulation"));

        body.Append(CreateCallout("The Two-Week Wait: During the period between ovulation or embryo transfer and pregnancy testing, anxiety often peaks. Use this time to practice gentle self-care, engage in light enjoyable activities, and remind yourself that you have done everything within your power. The outcome is not within your control — but your response to the waiting is."));

        // ==================== PART FOUR ====================
        body.Append(CreateHeading1("Part Four: Emotional Wellness and Your Support Circle", "_TocPart4"));

        body.Append(CreatePartIntro("The journey through fertility challenges touches every dimension of your being — physical, emotional, relational, and spiritual. This section offers companionship, practical coping strategies, and the wisdom of those who have walked this road before you."));

        // Chapter 10
        body.Append(CreateHeading2("Chapter 10: Honoring Your Emotional Journey"));

        body.Append(CreateChapterOpener("The path to motherhood at any age is emotional. At this season of life, it carries unique layers of complexity. Giving yourself permission to feel it all is not weakness — it is the foundation of resilience."));

        body.Append(CreateHeading3("The Grief No One Talks About"));

        body.Append(CreateParagraph("Women pursuing pregnancy during or after menopause often experience a form of disenfranchised grief — losses that society does not fully acknowledge or validate. You may be grieving the loss of time, the loss of the \"easy\" path to motherhood you imagined, the loss of using your own genetic material, or the loss of the young mother identity you once held."));

        body.Append(CreateParagraph("This grief is real, and it deserves attention. Allow yourself to mourn what has been lost even as you hope for what may still be. Holding both grief and hope simultaneously is not contradiction — it is the essence of courage."));

        body.Append(CreateQuote("\"Grief is just love with nowhere to go. Honor your grief, and it will eventually soften into wisdom.\""));

        body.Append(CreateHeading3("Navigating the Emotional Roller Coaster"));

        body.Append(CreateParagraph("Fertility treatment is inherently unpredictable. One month brings hope, the next brings disappointment. Hormonal medications amplify emotional responses. The financial investment adds pressure. Understanding this volatility can help you prepare and respond with self-compassion."));

        body.Append(CreateBulletParagraph("\u2022 Name your emotions as they arise — suppression only intensifies feelings"));
        body.Append(CreateBulletParagraph("\u2022 Create a ritual for processing disappointment — journaling, walking, talking with a trusted friend"));
        body.Append(CreateBulletParagraph("\u2022 Set boundaries with people who minimize your experience or offer unsolicited advice"));
        body.Append(CreateBulletParagraph("\u2022 Celebrate small victories along the way — a good response to medication, a healthy embryo"));
        body.Append(CreateBulletParagraph("\u2022 Consider working with a therapist who specializes in fertility challenges"));

        body.Append(CreateHeading3("The Partner Dimension"));

        body.Append(CreateParagraph("If you have a partner, this journey affects them too — though often in different ways. Men may feel pressure to \"fix\" the situation, helplessness at their limited role, or grief about using donor sperm if that becomes necessary. Open communication about each person's needs, fears, and hopes strengthens your bond during this challenging time."));

        body.Append(CreateBulletParagraph("\u2022 Schedule regular check-ins to share feelings without problem-solving"));
        body.Append(CreateBulletParagraph("\u2022 Allow your partner their own emotional response — it may differ from yours"));
        body.Append(CreateBulletParagraph("\u2022 Seek couples counseling if communication becomes strained"));
        body.Append(CreateBulletParagraph("\u2022 Find ways to nurture your relationship beyond fertility discussions"));

        // Chapter 11
        body.Append(CreateHeading2("Chapter 11: Building Your Support System"));

        body.Append(CreateChapterOpener("No woman should walk this path alone. Building a circle of support — professional, personal, and communal — can make the difference between barely surviving and genuinely thriving through this journey."));

        body.Append(CreateHeading3("Your Professional Support Team"));

        body.Append(CreateParagraph("Assemble a team that addresses all dimensions of your well-being:"));

        body.Append(CreateBulletParagraph("\u2022 Reproductive Endocrinologist — your medical director and fertility expert"));
        body.Append(CreateBulletParagraph("\u2022 Fertility Nurse Coordinator — your day-to-day guide through treatment"));
        body.Append(CreateBulletParagraph("\u2022 Reproductive Therapist — emotional support specialized to fertility challenges"));
        body.Append(CreateBulletParagraph("\u2022 Acupuncturist — stress reduction and potential fertility enhancement"));
        body.Append(CreateBulletParagraph("\u2022 Nutritionist — personalized dietary guidance for your specific needs"));
        body.Append(CreateBulletParagraph("\u2022 Obstetrician — for preconception counseling and eventual pregnancy care"));

        body.Append(CreateHeading3("Finding Your Tribe"));

        body.Append(CreateParagraph("Connecting with women who truly understand your experience is invaluable. Support groups — both in-person and online — provide a space to share openly without explanation or justification. Hearing others' stories normalizes your experience and offers practical insights you might not find elsewhere."));

        body.Append(CreateBulletParagraph("\u2022 Resolve (resolve.org) offers support groups across the United States"));
        body.Append(CreateBulletParagraph("\u2022 Fertility-focused Instagram and Facebook communities connect you globally"));
        body.Append(CreateBulletParagraph("\u2022 Many fertility clinics host their own support groups"));
        body.Append(CreateBulletParagraph("\u2022 Consider starting your own small group with women you meet along the way"));

        body.Append(CreateHeading3("Protecting Your Peace"));

        body.Append(CreateParagraph("Not everyone in your life will understand or support your journey, and that is okay. You have the right to set boundaries that protect your emotional well-being. This might mean limiting time with certain family members, declining baby showers when you are not up to attending, or being selective about who you tell details of your treatment."));

        body.Append(CreateCallout("Your journey is yours alone. You do not owe anyone an explanation for your choices, your timeline, or your emotional responses. Protect your energy fiercely — you need it for the work of becoming a mother."));

        // Chapter 12
        body.Append(CreateHeading2("Chapter 12: Stories of Hope and Possibility"));

        body.Append(CreateChapterOpener("These stories are shared with permission from women who have walked the path you now travel. Their journeys are unique, but the thread of resilience runs through them all."));

        body.Append(CreateHeading3("Maria's Story: Natural Conception at 44"));

        body.Append(CreateParagraph("\"I was told at 42 that my FSH was too high and I should consider donor eggs immediately. Something in me resisted giving up so quickly. I spent six months completely transforming my diet, started acupuncture, and worked with a functional medicine doctor to address my thyroid issues. At 44, I conceived naturally — completely unexpectedly, actually. My daughter is now three, and I am so grateful I trusted my instincts and gave my body a chance.\""));

        body.Append(CreateHeading3("Jennifer's Story: IVF with Own Eggs at 45"));

        body.Append(CreateParagraph("\"My first IVF cycle at 43 produced one embryo, which failed to implant. I was devastated. But my doctor adjusted my protocol, and the second cycle yielded two embryos. We transferred both, and one stuck. My son was born when I was 45. It took two years, three cycles, and every penny of our savings, but holding him erased every moment of pain.\""));

        body.Append(CreateHeading3("Sarah's Story: Donor Eggs at 51"));

        body.Append(CreateParagraph("\"After failed IVF with my own eggs at 47 and 48, I knew I needed to grieve that dream and embrace another. Choosing donor eggs felt like admitting defeat at first, but my therapist helped me see it differently. On my first donor egg transfer, I became pregnant with twins. I gave birth at 51, and I have never been happier. They are unequivocally my children — every sleepless night, every milestone, every moment of joy confirms what my heart always knew.\""));

        body.Append(CreateHeading3("Lisa's Story: The Journey to Acceptance"));

        body.Append(CreateParagraph("\"After three failed IVF cycles and two donor egg attempts, I reached a point where continuing felt more harmful than stopping. With tremendous support from my therapist and husband, I made the painful decision to end treatment. Two years later, we adopted a beautiful baby girl. She is the light of our lives. My journey to motherhood did not look like I imagined, but it led me exactly where I was meant to be.\""));

        body.Append(CreateCallout("Every story is different. Your story is still being written. Whether your path leads to pregnancy with your own eggs, donor eggs, adoption, or a different fulfillment of your nurturing spirit, your worth as a woman is not defined by your ability to conceive. You are already whole."));

        // ==================== CLOSING WORDS ====================
        body.Append(CreateHeading1("Closing Words: Your Journey Forward", "_TocClosing"));

        body.Append(CreateParagraph("As you close this book and step forward into whatever comes next, I want to leave you with a few final thoughts."));

        body.Append(CreateParagraph("First, trust your timing. The path to motherhood rarely follows our preferred schedule. Whether you conceive next month or next year, whether you use your own eggs or donor eggs, whether your child comes to you through birth or adoption — the love you will share is not diminished by the route taken to find each other."));

        body.Append(CreateParagraph("Second, take excellent care of yourself. Not because you are trying to earn a pregnancy, but because you deserve to feel vibrant, nourished, and whole regardless of the outcome. The habits you cultivate now — the healthy eating, the movement, the stress management, the sleep — will serve you for the rest of your life."));

        body.Append(CreateParagraph("Third, hold onto hope, but hold it lightly. Hope is not the enemy, even when it has been disappointed before. Let your hope be informed by science, grounded in reality, and flexible enough to adapt as your journey unfolds."));

        body.Append(CreateQuote("\"The strongest women are not those who show strength in front of us, but those who win battles we know nothing about.\""));

        body.Append(CreateParagraph("Finally, remember that you are not alone. Thousands of women have walked this path before you. Thousands walk it beside you now. Your desire for motherhood, your grief, your hope, your courage — these connect you to a sisterhood that spans all ages and circumstances."));

        body.Append(CreateParagraph("Wherever this journey leads you, may you find peace in knowing you did not let your dream go without a fight. May you find joy in the life you build, whatever form it takes. And may you always remember that your value as a woman extends far beyond your fertility — you are worthy of love, happiness, and fulfillment exactly as you are."));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "600", After = "200" }),
            new Run(new RunProperties(new FontSize { Val = "28" }, new Italic(), new Color { Val = Colors.Secondary }),
                new Text("\u2736 \u2736 \u2736"))));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "100" }),
            new Run(new RunProperties(new FontSize { Val = "22" }, new Italic(), new Color { Val = Colors.Mid }),
                new Text("With love and hope for your journey,"))));

        // Content section break
        body.Append(new Paragraph(new ParagraphProperties(new SectionProperties(
            new HeaderReference { Type = HeaderFooterValues.Default, Id = headerId },
            new FooterReference { Type = HeaderFooterValues.Default, Id = footerId },
            new PageSize { Width = (UInt32Value)(uint)A4W, Height = (UInt32Value)(uint)A4H },
            new PageMargin { Top = 1800, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720 }))));
    }

    // ========================================================================
    // Backcover
    // ========================================================================
    private static void AddBackcoverSection(Body body, string backBgId, ref uint prId)
    {
        body.Append(new Paragraph(new Run(CreateFloatingBackground(backBgId, prId++, "BackBg"))));

        body.Append(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "6000" }), new Run()));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "400" }),
            new Run(new RunProperties(new FontSize { Val = "36" }, new Bold(), new Italic(),
                new Color { Val = Colors.Primary },
                new RunFonts { Ascii = "Georgia", HighAnsi = "Georgia" }),
                new Text("Embracing Possibility"))));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "600" }),
            new Run(new RunProperties(new FontSize { Val = "22" }, new Italic(), new Color { Val = Colors.Mid }),
                new Text("Honest Science. Real Hope. Practical Wisdom."))));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new Indentation { Left = "1200", Right = "1200" },
                new SpacingBetweenLines { After = "600" }),
            new Run(new RunProperties(new FontSize { Val = "21" }, new Color { Val = Colors.Dark }),
                new Text("This comprehensive guide offers evidence-based information about fertility during perimenopause, IVF and donor egg options, nutrition and lifestyle strategies, and emotional support for the journey to motherhood after 40."))));

        // Decorative
        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "600" }),
            new Run(new RunProperties(new Color { Val = Colors.Accent }, new FontSize { Val = "24" }),
                new Text("\u2014\u2014\u2014  \u2736  \u2014\u2014\u2014"))));

        body.Append(new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
            new Run(new RunProperties(new FontSize { Val = "20" }, new Color { Val = Colors.Light }),
                new Text("Your journey to motherhood deserves honest guidance and compassionate support."))));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "200" }),
            new Run(new RunProperties(new FontSize { Val = "18" }, new Color { Val = Colors.Light }),
                new Text("\u00a9 2025  |  All rights reserved"))));

        // Final section
        body.Append(new SectionProperties(
            new PageSize { Width = (UInt32Value)(uint)A4W, Height = (UInt32Value)(uint)A4H },
            new PageMargin { Top = 0, Right = 0, Bottom = 0, Left = 0, Header = 0, Footer = 0 }));
    }

    // ========================================================================
    // Factory helpers
    // ========================================================================
    private static int _bookmarkId = 0;

    private static Paragraph CreateHeading1(string text, string bookmarkName)
    {
        int id = ++_bookmarkId;
        return new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new BookmarkStart { Id = id.ToString(), Name = bookmarkName },
            new Run(new Text(text)),
            new BookmarkEnd { Id = id.ToString() });
    }

    private static Paragraph CreateHeading2(string text)
    {
        return new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading2" }),
            new Run(new Text(text)));
    }

    private static Paragraph CreateHeading3(string text)
    {
        return new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading3" }),
            new Run(new Text(text)));
    }

    private static Paragraph CreateParagraph(string text)
    {
        return new Paragraph(new Run(new Text(text)));
    }

    private static Paragraph CreateQuote(string text)
    {
        return new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Quote" }),
            new Run(new Text(text)));
    }

    private static Paragraph CreateBulletParagraph(string text)
    {
        return new Paragraph(
            new ParagraphProperties(
                new Indentation { Left = "360" },
                new SpacingBetweenLines { After = "120" }),
            new Run(new Text(text)));
    }

    private static Paragraph CreateChapterOpener(string text)
    {
        return new Paragraph(
            new ParagraphProperties(
                new Indentation { Left = "360", Right = "360" },
                new SpacingBetweenLines { Before = "120", After = "300" },
                new ParagraphBorders(
                    new LeftBorder { Val = BorderValues.Single, Size = 12, Color = Colors.Secondary, Space = 12 })),
            new Run(new RunProperties(new Italic(), new FontSize { Val = "22" }, new Color { Val = Colors.Mid }),
                new Text(text)));
    }

    private static Paragraph CreatePartIntro(string text)
    {
        return new Paragraph(
            new ParagraphProperties(
                new Indentation { Left = "720", Right = "720" },
                new SpacingBetweenLines { Before = "100", After = "400" }),
            new Run(new RunProperties(new Italic(), new FontSize { Val = "24" }, new Color { Val = Colors.Mid }),
                new Text(text)));
    }

    private static Paragraph CreateCaption(string text)
    {
        return new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }),
            new Run(new Text(text)));
    }

    private static Paragraph CreateCallout(string text)
    {
        return new Paragraph(
            new ParagraphProperties(
                new Shading { Val = ShadingPatternValues.Clear, Fill = Colors.CalloutBg },
                new Indentation { Left = "360", Right = "360" },
                new SpacingBetweenLines { Before = "200", After = "200" },
                new ParagraphBorders(
                    new LeftBorder { Val = BorderValues.Single, Size = 18, Color = Colors.Secondary, Space = 8 })),
            new Run(new RunProperties(new Italic(), new FontSize { Val = "21" }, new Color { Val = Colors.Dark }),
                new Text(text)));
    }

    // ========================================================================
    // Tables
    // ========================================================================
    private static Table CreateFertilityTable()
    {
        var tbl = new Table(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 12, Color = Colors.Primary },
                new BottomBorder { Val = BorderValues.Single, Size = 12, Color = Colors.Primary },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = Colors.Border })),
            new TableGrid(new GridColumn { Width = "2500" }, new GridColumn { Width = "2500" }, new GridColumn { Width = "2500" }, new GridColumn { Width = "2500" }));
        var w1 = new[] { "2500", "2500", "2500", "2500" };
        tbl.Append(CreateTableRow(true, w1, new[] { "Age Group", "Chance per Cycle", "Chance in 1 Year", "IVF Success Rate" }));
        tbl.Append(CreateTableRow(false, w1, new[] { "Under 30", "20-25%", "85%", "40-50%" }));
        tbl.Append(CreateTableRow(false, w1, new[] { "30-34", "15-20%", "75%", "35-40%" }));
        tbl.Append(CreateTableRow(false, w1, new[] { "35-37", "10-15%", "65%", "25-30%" }));
        tbl.Append(CreateTableRow(false, w1, new[] { "38-40", "5-10%", "40%", "15-20%" }));
        tbl.Append(CreateTableRow(false, w1, new[] { "41-42", "3-5%", "20%", "8-12%" }));
        tbl.Append(CreateTableRow(false, w1, new[] { "43-44", "1-3%", "5%", "3-5%" }));
        tbl.Append(CreateTableRow(false, w1, new[] { "45+", "<1%", "<3%", "1-2%" }));
        return tbl;
    }

    private static Table CreateTestsTable()
    {
        var tbl = new Table(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 12, Color = Colors.Primary },
                new BottomBorder { Val = BorderValues.Single, Size = 12, Color = Colors.Primary },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = Colors.Border })),
            new TableGrid(new GridColumn { Width = "2500" }, new GridColumn { Width = "7500" }));
        var w2 = new[] { "2500", "7500" };
        tbl.Append(CreateTableRow(true, w2, new[] { "Test Name", "What It Measures" }));
        tbl.Append(CreateTableRow(false, w2, new[] { "AMH (Anti-Mullerian Hormone)", "Egg quantity remaining in ovaries" }));
        tbl.Append(CreateTableRow(false, w2, new[] { "FSH (Day 3)", "Ovarian reserve and egg quality signals" }));
        tbl.Append(CreateTableRow(false, w2, new[] { "Estradiol (Day 3)", "Baseline estrogen level" }));
        tbl.Append(CreateTableRow(false, w2, new[] { "Antral Follicle Count", "Visible egg-containing follicles via ultrasound" }));
        tbl.Append(CreateTableRow(false, w2, new[] { "Thyroid Panel", "Thyroid function affecting fertility" }));
        tbl.Append(CreateTableRow(false, w2, new[] { "Prolactin", "Hormone that can suppress ovulation" }));
        tbl.Append(CreateTableRow(false, w2, new[] { "Vitamin D", "Essential for fertility and pregnancy" }));
        return tbl;
    }

    private static Table CreateQuestionsTable()
    {
        var tbl = new Table(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 12, Color = Colors.Primary },
                new BottomBorder { Val = BorderValues.Single, Size = 12, Color = Colors.Primary },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = Colors.Border })),
            new TableGrid(new GridColumn { Width = "5000" }, new GridColumn { Width = "5000" }));
        var w3 = new[] { "5000", "5000" };
        tbl.Append(CreateTableRow(true, w3, new[] { "Question", "Why It Matters" }));
        tbl.Append(CreateTableRow(false, w3, new[] { "What are my realistic chances?", "Sets honest expectations from the start" }));
        tbl.Append(CreateTableRow(false, w3, new[] { "Would you recommend donor eggs?", "Determines if you're a candidate" }));
        tbl.Append(CreateTableRow(false, w3, new[] { "What protocol do you recommend?", "Shows the treatment plan for your case" }));
        tbl.Append(CreateTableRow(false, w3, new[] { "What are your success rates?", "Compares clinic performance data" }));
        tbl.Append(CreateTableRow(false, w3, new[] { "What is the total cost?", "Prepares you financially" }));
        return tbl;
    }

    private static Table CreateIVFSuccessTable()
    {
        var tbl = new Table(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 12, Color = Colors.Primary },
                new BottomBorder { Val = BorderValues.Single, Size = 12, Color = Colors.Primary },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = Colors.Border })),
            new TableGrid(new GridColumn { Width = "3333" }, new GridColumn { Width = "3333" }, new GridColumn { Width = "3334" }));
        var w4 = new[] { "3333", "3333", "3334" };
        tbl.Append(CreateTableRow(true, w4, new[] { "Age Group", "Live Birth per Cycle", "Cumulative (3 cycles)" }));
        tbl.Append(CreateTableRow(false, w4, new[] { "Under 35", "40-50%", "70-80%" }));
        tbl.Append(CreateTableRow(false, w4, new[] { "35-37", "35-40%", "60-70%" }));
        tbl.Append(CreateTableRow(false, w4, new[] { "38-40", "25-30%", "45-55%" }));
        tbl.Append(CreateTableRow(false, w4, new[] { "41-42", "15-20%", "30-40%" }));
        tbl.Append(CreateTableRow(false, w4, new[] { "43-44", "5-10%", "15-20%" }));
        tbl.Append(CreateTableRow(false, w4, new[] { "45+", "1-3%", "3-8%" }));
        return tbl;
    }

    private static Table CreateCostTable()
    {
        var tbl = new Table(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 12, Color = Colors.Primary },
                new BottomBorder { Val = BorderValues.Single, Size = 12, Color = Colors.Primary },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = Colors.Border })),
            new TableGrid(new GridColumn { Width = "6000" }, new GridColumn { Width = "4000" }));
        var w5 = new[] { "6000", "4000" };
        tbl.Append(CreateTableRow(true, w5, new[] { "Service", "Estimated Cost (USD)" }));
        tbl.Append(CreateTableRow(false, w5, new[] { "Donor compensation & agency fees", "$8,000 - $15,000" }));
        tbl.Append(CreateTableRow(false, w5, new[] { "Donor medical expenses", "$3,000 - $6,000" }));
        tbl.Append(CreateTableRow(false, w5, new[] { "IVF cycle with donor eggs", "$15,000 - $25,000" }));
        tbl.Append(CreateTableRow(false, w5, new[] { "Medications for recipient", "$1,000 - $3,000" }));
        tbl.Append(CreateTableRow(false, w5, new[] { "Genetic testing (PGT) of embryos", "$3,000 - $6,000" }));
        tbl.Append(CreateTableRow(false, w5, new[] { "Frozen embryo transfer (if needed)", "$3,000 - $5,000" }));
        tbl.Append(CreateTableRow(false, w5, new[] { "Legal fees", "$500 - $2,000" }));
        tbl.Append(CreateTableRow(false, w5, new[] { "Total estimated range", "$33,500 - $62,000" }));
        return tbl;
    }

    private static Table CreateNutritionTable()
    {
        var tbl = new Table(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 12, Color = Colors.Primary },
                new BottomBorder { Val = BorderValues.Single, Size = 12, Color = Colors.Primary },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = Colors.Border })),
            new TableGrid(new GridColumn { Width = "2500" }, new GridColumn { Width = "5000" }, new GridColumn { Width = "2500" }));
        var w6 = new[] { "2500", "5000", "2500" };
        tbl.Append(CreateTableRow(true, w6, new[] { "Category", "Food Examples", "Serving" }));
        tbl.Append(CreateTableRow(false, w6, new[] { "Vegetables (1/2 plate)", "Spinach, broccoli, peppers, beets", "2-3 cups" }));
        tbl.Append(CreateTableRow(false, w6, new[] { "Protein (1/4 plate)", "Salmon, chicken, lentils, eggs", "4-6 oz" }));
        tbl.Append(CreateTableRow(false, w6, new[] { "Whole Grains (1/4 plate)", "Quinoa, brown rice, oats", "1/2-1 cup" }));
        tbl.Append(CreateTableRow(false, w6, new[] { "Healthy Fats", "Avocado, olive oil, nuts, seeds", "1-2 tbsp" }));
        tbl.Append(CreateTableRow(false, w6, new[] { "Fermented Foods", "Yogurt, kefir, sauerkraut", "1/4-1/2 cup" }));
        tbl.Append(CreateTableRow(false, w6, new[] { "Fruit", "Berries, citrus, pomegranate", "1 cup" }));
        return tbl;
    }

    private static Table CreateSampleMealsTable()
    {
        var tbl = new Table(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 12, Color = Colors.Primary },
                new BottomBorder { Val = BorderValues.Single, Size = 12, Color = Colors.Primary },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = Colors.Border })),
            new TableGrid(new GridColumn { Width = "2000" }, new GridColumn { Width = "4000" }, new GridColumn { Width = "4000" }));
        var w7 = new[] { "2000", "4000", "4000" };
        tbl.Append(CreateTableRow(true, w7, new[] { "Meal", "Option A", "Option B" }));
        tbl.Append(CreateTableRow(false, w7, new[] { "Breakfast", "Quinoa bowl with berries, walnuts, and Greek yogurt", "Veggie omelet with avocado and whole grain toast" }));
        tbl.Append(CreateTableRow(false, w7, new[] { "Lunch", "Grilled salmon salad with mixed greens and olive oil", "Lentil soup with side salad and flaxseed crackers" }));
        tbl.Append(CreateTableRow(false, w7, new[] { "Dinner", "Baked chicken with roasted vegetables and brown rice", "Stir-fried tofu with broccoli, peppers, and quinoa" }));
        tbl.Append(CreateTableRow(false, w7, new[] { "Snacks", "Apple slices with almond butter", "Mixed nuts and dark chocolate (70% cacao)" }));
        return tbl;
    }

    private static TableRow CreateTableRow(bool hdr, string[] widths, string[] cells)
    {
        var row = new TableRow();
        if (hdr) row.Append(new TableRowProperties(new TableHeader()));
        for (int i = 0; i < cells.Length; i++)
        {
            var w = i < widths.Length ? widths[i] : "2500";
            var tcp = new TableCellProperties(new TableCellWidth { Width = w, Type = TableWidthUnitValues.Dxa });
            if (hdr) tcp.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = Colors.TableHeader });
            var rpr = new RunProperties(new FontSize { Val = "20" }, new Color { Val = hdr ? Colors.Dark : Colors.Mid });
            if (hdr) rpr.Append(new Bold());
            row.Append(new TableCell(tcp, new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Center },
                    new SpacingBetweenLines { Before = "60", After = "60" }),
                new Run(rpr, new Text(cells[i])))));
        }
        return row;
    }

    // ========================================================================
    // Image helpers
    // ========================================================================
    private static string AddImage(MainDocumentPart mp, string path)
    {
        var ip = mp.AddImagePart(ImagePartType.Png);
        using var fs = new FileStream(path, FileMode.Open);
        ip.FeedData(fs); return mp.GetIdOfPart(ip);
    }

    private static Drawing CreateFloatingBackground(string imgId, uint prId, string name)
    {
        return new Drawing(new DW.Anchor(
            new DW.SimplePosition { X = 0, Y = 0 },
            new DW.HorizontalPosition(new DW.PositionOffset("0")) { RelativeFrom = DW.HorizontalRelativePositionValues.Page },
            new DW.VerticalPosition(new DW.PositionOffset("0")) { RelativeFrom = DW.VerticalRelativePositionValues.Page },
            new DW.Extent { Cx = A4WE, Cy = A4HE },
            new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            new DW.WrapNone(),
            new DW.DocProperties { Id = prId, Name = name },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(new A.GraphicData(
                new PIC.Picture(
                    new PIC.NonVisualPictureProperties(
                        new PIC.NonVisualDrawingProperties { Id = 0, Name = $"{name}.png" },
                        new PIC.NonVisualPictureDrawingProperties()),
                    new PIC.BlipFill(new A.Blip { Embed = imgId }, new A.Stretch(new A.FillRectangle())),
                    new PIC.ShapeProperties(
                        new A.Transform2D(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = A4WE, Cy = A4HE }),
                        new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle })))
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
        { DistanceFromTop = 0, DistanceFromBottom = 0, DistanceFromLeft = 0, DistanceFromRight = 0,
          SimplePos = false, RelativeHeight = 251658240, BehindDoc = true,
          Locked = false, LayoutInCell = true, AllowOverlap = true });
    }

    // ========================================================================
    // Settings
    // ========================================================================
    private static void SetUpdateFieldsOnOpen(MainDocumentPart mp)
    {
        var sp = mp.DocumentSettingsPart ?? mp.AddNewPart<DocumentSettingsPart>();
        sp.Settings = new Settings(new UpdateFieldsOnOpen { Val = true }, new DisplayBackgroundShape());
    }

    private static void AddNumbering(MainDocumentPart mp)
    {
        var np = mp.AddNewPart<NumberingDefinitionsPart>();
        np.Numbering = new Numbering(
            new AbstractNum(new Level(
                new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = "%1." },
                new LevelJustification { Val = LevelJustificationValues.Left },
                new ParagraphProperties(new Indentation { Left = "720", Hanging = "360" })
            ) { LevelIndex = 0 }) { AbstractNumberId = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
    }
}

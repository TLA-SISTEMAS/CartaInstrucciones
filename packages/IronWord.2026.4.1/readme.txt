IronWord: The C# Word Library
===================================================================================
IronWord is a C# library for creating, reading, and editing Word documents in .NET.
===================================================================================

Build Status: Passing
NuGet Version: 2026.4.1 [https://www.nuget.org/packages/IronWord]
License: Commercial

✨ KEY FEATURES
  * ➕ Create Words: Generate and populate new Word (docx) files from scratch.

  * ✍️ Edit Words: Programmatically find and replace text, insert images and shapes, and extract content for data processing.

  * 🎨 Style Words: Apply rich formatting and styling to any document element including texts and images.


📦 INSTALLATION

To install IronWord via the NuGet Package Manager Console, run the following command:

    PM> Install-Package IronWord


🚀 QUICK START

Creating a Word document is simple. After installing the package, you can create a .docx file with just a few lines of code:

    using IronWord;

    // Create textrun
    TextContent textRun = new TextContent("Hello World from IronWord!");

    Paragraph paragraph = new Paragraph();
    paragraph.AddChild(textRun);

    // Create a new Word document
    WordDocument doc = new WordDocument(paragraph);

    // Export docx
    doc.SaveAs("document.docx");

For more detailed examples, please visit our website's quick-start guide: https://ironsoftware.com/csharp/word/docs/


📚 DOCUMENTATION

Our documentation is structured to help you learn effectively, whether you're a beginner or an expert.

  * 🎓 Tutorials: Step-by-step guides to help you build your first Word application.
    * C# Word Tutorials: https://ironsoftware.com/csharp/word/tutorials/document-element/

  * 🛠️ How-To Guides: Practical, goal-oriented instructions to solve specific problems.
    * Code Examples: https://ironsoftware.com/csharp/word/examples/create-empty-word/
    * How-To Articles: https://ironsoftware.com/csharp/word/docs/

  * 📖 Reference: Detailed technical descriptions of the API and its components.
    * API Reference: https://ironsoftware.com/csharp/word/object-reference/api/

  * 🤔 Demos: Demonstrations that showcase how IronWord works the way it does.
    * Book a Live Demo: https://ironsoftware.com/csharp/word/examples/create-empty-word/#booking-demo


🤝 GETTING HELP & SUPPORT

Have a question or running into an issue?

  * Email Support: Reach out to our team directly at support@ironsoftware.com.
  * Report a Bug: https://ironsoftware.com/ticket-submission/
  * Community: https://ironsoftware.com/company/iron-slack-community/


💻 COMPATIBILITY

  * Languages: C#, F#, and VB.NET
  * Platforms: .NET 10, .NET 9, .NET 8, .NET 7, .NET 6, .NET 5, Core 2x & 3x, Standard 2
  * Framework: .NET Framework 4.6.2 (and above)
  * App Models: Console, Web, and Desktop Apps
  * Operating Systems: Windows, macOS, Linux (Debian, CentOS, Ubuntu)
  * Cloud & Containerization: Azure, AWS, Docker
  * IDEs: Microsoft Visual Studio, Jetbrains ReSharper, JetBrains Rider


📝 LICENSE

IronWord is a commercially licensed product.
  * Trial License: https://ironsoftware.com/csharp/word/licensing/new/#trial-license
  * Purchase a License: https://ironsoftware.com/csharp/word/licensing/new/

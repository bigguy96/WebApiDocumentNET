using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.OpenApi.Readers;

namespace ApiDocumentationWithSwagger
{
    // Model to represent an API endpoint
    public class ApiEndpoint
    {
        public string? Method { get; set; }
        public string? Path { get; set; }
        public string? Description { get; set; }
        public List<string>? Parameters { get; set; }
        public string? Response { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string myDocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string outputDocPath = Path.Combine(myDocumentsPath, "ApiDocumentation.docx");
            string swaggerJsonPath = Path.Combine(myDocumentsPath, "swagger.json");

            // Parse Swagger JSON and extract endpoints
            var endpoints = ParseSwaggerJson(swaggerJsonPath);

            // Generate Word document with coloring
            CreateApiDocumentation(endpoints, outputDocPath);

            Console.WriteLine($"Documentation generated at: {outputDocPath}");
        }

        static List<ApiEndpoint> ParseSwaggerJson(string swaggerJsonPath)
        {
            var endpoints = new List<ApiEndpoint>();

            // Read the Swagger JSON file
            var json = File.ReadAllText(swaggerJsonPath);
            var openApiDoc = new OpenApiStringReader().Read(json, out var diagnostic);

            // if (diagnostic.Errors.Any())
            // {
            //     Console.WriteLine("Errors parsing Swagger JSON:");
            //     foreach (var error in diagnostic.Errors)
            //         Console.WriteLine(error.Message);
            //     return endpoints;
            // }

            // Extract API endpoints from the Swagger document
            foreach (var path in openApiDoc.Paths)
            {
                foreach (var operation in path.Value.Operations)
                {
                    var endpoint = new ApiEndpoint
                    {
                        Method = operation.Key.ToString().ToUpper(),
                        Path = path.Key,
                        Description = operation.Value.Summary ?? operation.Value.Description ?? "No description provided",
                        Parameters = new List<string>(),
                        Response = operation.Value.Responses.TryGetValue("200", out var response)
                            ? $"200 OK: {response.Description}"
                            : "No response info"
                    };

                    // Extract parameters
                    if (operation.Value.Parameters != null)
                    {
                        foreach (var param in operation.Value.Parameters)
                        {
                            endpoint.Parameters.Add($"{param.Name} ({param.Schema?.Type ?? "unknown"}): {param.Description ?? "No description"}");
                        }
                    }

                    // If the operation has a request body (e.g., for POST)
                    if (operation.Value.RequestBody != null)
                    {
                        var requestBody = operation.Value.RequestBody;
                        var content = requestBody.Content.FirstOrDefault().Value?.Schema;
                        endpoint.Parameters.Add($"Request Body: {requestBody.Description ?? "No description"}");
                    }

                    endpoints.Add(endpoint);
                }
            }

            return endpoints;
        }

        static void CreateApiDocumentation(List<ApiEndpoint> endpoints, string filePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = new Body();

                // Add a colored title
                Paragraph title = new Paragraph(
                    new Run(
                        new Text("API Documentation")
                    )
                    {
                        RunProperties = new RunProperties(
                            new Bold(),
                            new FontSize { Val = "32" }, // Font size 16pt
                            new Color { Val = "2E75B5" } // Blue color
                        )
                    }
                );
                body.Append(title);

                // Add a brief introduction
                Paragraph intro = new Paragraph(
                    new Run(
                        new Text("This document provides details of the API endpoints, including methods, parameters, and responses.")
                    )
                    {
                        RunProperties = new RunProperties(
                            new Color { Val = "404040" } // Dark gray color
                        )
                    }
                );
                body.Append(intro);

                // Add each endpoint
                foreach (var endpoint in endpoints)
                {
                    // Determine color based on HTTP method
                    string headingColor = endpoint.Method switch
                    {
                        "GET" => "00FF00",    // Green for GET
                        "POST" => "2E75B5",   // Blue for POST
                        "PUT" => "ED7D31",    // Orange for PUT
                        "DELETE" => "FF0000", // Red for DELETE
                        _ => "808080"         // Gray for other methods (e.g., PATCH, OPTIONS)
                    };

                    // Endpoint heading (e.g., "GET /api/users") with method-based color
                    Paragraph endpointTitle = new Paragraph(
                        new Run(
                            new Text($"{endpoint.Method} {endpoint.Path}")
                        )
                        {
                            RunProperties = new RunProperties(
                                new Bold(),
                                new FontSize { Val = "24" }, // Font size 12pt
                                new Color { Val = headingColor } // Method-based color
                            )
                        }
                    );
                    body.Append(endpointTitle);

                    // Description
                    Paragraph desc = new Paragraph(
                        new Run(
                            new Text($"Description: {endpoint.Description}")
                        )
                        {
                            RunProperties = new RunProperties(
                                new Color { Val = "000000" } // Black color
                            )
                        }
                    );
                    body.Append(desc);

                    // Parameters heading with blue color
                    Paragraph paramTitle = new Paragraph(
                        new Run(
                            new Text("Parameters:")
                        )
                        {
                            RunProperties = new RunProperties(
                                new Bold(),
                                new Color { Val = "2E75B5" } // Blue color
                            )
                        }
                    );
                    body.Append(paramTitle);

                    // Parameters list
                    if (endpoint.Parameters != null && endpoint.Parameters.Count > 0)
                    {
                        foreach (var param in endpoint.Parameters)
                        {
                            Paragraph paramItem = new Paragraph(
                                new Run(
                                    new Text($"  - {param}")
                                )
                                {
                                    RunProperties = new RunProperties(
                                        new Color { Val = "000000" } // Black color
                                    )
                                }
                            );
                            body.Append(paramItem);
                        }
                    }
                    else
                    {
                        Paragraph noParams = new Paragraph(
                            new Run(
                                new Text("  - None")
                            )
                            {
                                RunProperties = new RunProperties(
                                    new Color { Val = "000000" }
                                )
                            }
                        );
                        body.Append(noParams);
                    }

                    // Response with purple color
                    Paragraph response = new Paragraph(
                        new Run(
                            new Text($"Response: {endpoint.Response}")
                        )
                        {
                            RunProperties = new RunProperties(
                                new Color { Val = "7030A0" } // Purple color
                            )
                        }
                    );
                    body.Append(response);

                    // Add spacing
                    body.Append(new Paragraph(new Run(new Text(""))));
                }

                // Append the body to the document
                mainPart.Document.Append(body);
                mainPart.Document.Save();
            }
        }
    }
}
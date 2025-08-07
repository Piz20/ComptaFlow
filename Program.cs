using ComptaFlow.Hubs;

var builder = WebApplication.CreateBuilder(args);

// 📦 Ajout des services nécessaires
builder.Services.AddControllers();  // ← Nécessaire pour MapControllers()
builder.Services.AddSignalR();

var app = builder.Build();

// 🚦 Configuration des middlewares
app.UseHttpsRedirection();
app.UseRouting();
app.UseAuthorization();
app.UseStaticFiles();  // Pour servir des fichiers statiques (HTML, JS, CSS)
// 🌐 Mapping des endpoints
app.MapControllers();                   // Pour les API
app.MapHub<ComptaHub>("/comptaHub");    // Pour SignalR

app.Run();

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorPages();

// Session storage
builder.Services.AddDistributedMemoryCache();

builder.Services.AddSession(options =>
{
    // Session expires after 30 minutes of inactivity
    options.IdleTimeout = TimeSpan.FromMinutes(30);

    // Make the session cookie HTTP-only
    options.Cookie.HttpOnly = true;

    // Required for session to work
    options.Cookie.IsEssential = true;
});

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();

app.UseStaticFiles();

app.UseRouting();

// Enable Session
app.UseSession();

app.UseAuthorization();

app.MapRazorPages();

app.Run();
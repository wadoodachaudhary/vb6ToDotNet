// Bootstrap utility: produce an ASP.NET Identity V3 hash for any
// (UserID, plaintext) pair and emit a ready-to-run SQL UPDATE that
// populates User_Manager.User_Password2 for that user.
//
// Use this to seed credentials for users whose User_Password value isn't
// in MyCrypt format (e.g. ADMIN / SYSTEM after the security review found
// non-MyCrypt stored values) — without it, those users can never satisfy
// opt 2's EnforcePasswordCheck verifier and so opt 4's auto-migration
// path never runs.
//
// Output is a single UPDATE statement parameterized to TOUCH ONLY
// User_Password2 — User_Password is never modified. Copy the output and
// paste into sqlcmd or your DB client.
//
// Usage:
//   dotnet run --project /Users/wadood/projects/VBToCSharp/HomeFront/tools/hash-password -- <UserID> <plaintext>
// Example:
//   dotnet run --project ... -- ADMIN admin
//
// The hash IS salt-bound to the UserID (the hasher takes user as the
// first arg so per-user pepper applies). Re-running with the same
// plaintext produces a different hash each time because the salt is
// randomized — that's normal.

using Microsoft.AspNetCore.Identity;

if (args.Length != 2) {
    Console.Error.WriteLine("usage: hash-password <UserID> <plaintext>");
    Console.Error.WriteLine("example: hash-password ADMIN admin");
    return 2;
}

var userId    = args[0];
var plaintext = args[1];

var hasher = new PasswordHasher<string>();
var hash   = hasher.HashPassword(userId, plaintext);

Console.WriteLine("-- ASP.NET Identity V3 hash for user " + userId);
Console.WriteLine("-- Run this in sqlcmd / SSMS / Azure Data Studio:");
Console.WriteLine();
Console.WriteLine("USE HOMEFRONTSQL;");
Console.WriteLine("UPDATE dbo.User_Manager");
Console.WriteLine($"   SET User_Password2 = '{hash}'");
Console.WriteLine($" WHERE User_ID = '{userId}';");
Console.WriteLine();
Console.WriteLine("-- Verify:");
Console.WriteLine($"SELECT User_ID, LEFT(User_Password2, 12) AS prefix, LEN(User_Password2) AS hash_len");
Console.WriteLine($"  FROM dbo.User_Manager WHERE User_ID = '{userId}';");
return 0;

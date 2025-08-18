# M365 User Provisioning Tool - Helpful Prompts for 1st Line Support

## Quick Start Prompts

### "Help me create a new user account"
> I need to create a new M365 user account for [Name]. Can you guide me through the process using the User Creation tab?

### "Show me how to do bulk user import"
> I have a CSV file with 20 new employees. Can you help me use the Bulk Import feature to create all their accounts at once?

### "The tool won't connect to M365"
> I'm having trouble connecting to Microsoft 365. Can you help me troubleshoot the authentication issues?

### "🔄 How do I switch between different tenants?"
> I need to switch from one Microsoft 365 tenant to another. Can you show me how to use the Switch Tenant feature?

### "🏢 MSP workflow for multiple clients"
> I'm a managed service provider working with multiple client tenants. What's the best workflow for switching between different client environments?

## User Creation Prompts

### Single User Creation
> Create a new user account with these details:
> - Name: [First Last]
> - Department: [Department]
> - Job Title: [Title]
> - Office: [Location]
> - Manager: [Manager Email]
> - License: [License Type]
> - Groups: [Group Names]

### Password Management
> Generate a secure password for the new user and ensure they must change it on first login.

### License Assignment
> Show me what licenses are available in our tenant and help me assign the right one for a [Job Role] in [Department].

## Bulk Import Prompts

### CSV Template Help
> I need to create 15 new users from an Excel file. Can you help me format the data correctly using the CSV template?

### Bulk Import Process
> Walk me through the bulk import process step by step. I have the CSV ready and need to create the accounts safely.

### Error Handling
> Some users failed during bulk import. Can you help me check the Activity Log and fix the errors?

## 🔄 Switch Tenant Prompts

### Basic Tenant Switching
> I need to switch from our main tenant to a different Microsoft 365 organization. Can you walk me through the Switch Tenant process step by step?

### MSP Client Management
> I'm working as an MSP and need to switch between different client tenants throughout the day. What's the most efficient workflow for this?

### Tenant Switch Troubleshooting
> I clicked Switch Tenant but I'm still seeing data from the previous tenant. Can you help me troubleshoot this issue?

### Authentication Issues After Switch
> After switching tenants, the browser keeps logging me into the old tenant automatically. How do I force authentication to a different tenant?

### Data Verification After Switch
> How can I verify that I've completely switched to a new tenant and that no old data is cached or visible?

### Emergency Tenant Access
> I need to quickly switch to a different tenant for an urgent issue. What's the fastest way to disconnect and reconnect?

### Cache Clearing Issues
> I switched tenants but I'm still seeing shared mailboxes and distribution lists from the previous tenant. How do I clear all cached data?

### Testing Tenant Switches
> I want to test the tenant switching functionality before using it in production. Can you help me validate that it's working correctly?

### Multi-Domain Environments
> Our organization has multiple M365 tenants with different domains. How do I ensure I'm connecting to the right tenant when switching?

## Troubleshooting Prompts

### Connection Issues
> The tool shows "Authentication failed" when I try to connect. Can you help me troubleshoot this?

### Module Problems
> I'm getting errors about missing PowerShell modules. Can you help me install the prerequisites?

### PowerShell Version
> How do I check if I have PowerShell 7 installed? The tool says I need it but I'm not sure what version I have.

### User Creation Failures
> I tried to create a user but got an error saying "[Error Message]". What does this mean and how do I fix it?

## Multi-Tenant Prompts

### Switching Tenants
> I need to switch from our main tenant to our development tenant. Can you help me use the Switch Tenant feature?

### Tenant Discovery
> After connecting to a new tenant, the tool is still loading tenant data. Is this normal and what should I see?

### MSP Workflows
> I'm a managed service provider working with multiple clients. Can you show me the best way to switch between different customer tenants?

## Advanced Usage Prompts

### CSV Template Customization
> Can you help me modify the CSV template to include additional fields for our organization's specific requirements?

### Automation Integration
> I want to integrate this tool with our ticketing system. Can you show me how to use the command-line options?

### Audit and Compliance
> I need to export the activity logs for our monthly audit. Can you show me how to access and export this data?

### Testing Mode
> Before creating real users in production, can you help me test the process safely without making actual changes?

## Error Resolution Prompts

### "Username already exists"
> I'm getting "Username already exists" errors. Can you help me check what usernames are available and suggest alternatives?

### "Insufficient licenses"
> The tool says there aren't enough licenses available. How do I check license usage and availability?

### "Invalid email format"
> Some of my bulk import entries are failing with email format errors. Can you help me validate the email addresses?

### "Manager not found"
> I'm trying to assign a manager but getting "Manager not found" errors. How do I find the correct manager email address?

## Group and License Management Prompts

### Group Assignment
> Show me all available security groups and distribution lists in our tenant so I can assign the new user to the right groups.

### License Types
> What's the difference between the available license types in our tenant? I need to assign the right license for a [Job Role].

### Distribution Lists
> How do I add the new user to our department's distribution list and the company-wide announcements list?

## Security and Compliance Prompts

### Password Policies
> What password requirements does the tool use when generating passwords? Can I customize these for our security policy?

### Audit Trail
> I need to show my manager what user accounts were created this month. Can you help me generate an audit report?

### Permission Requirements
> What minimum permissions do I need in M365 to use this tool? I want to request the least privileges necessary.

### Data Protection
> Does this tool store any sensitive data locally? I need to ensure we're compliant with our data protection policies.

## Daily Workflow Prompts

### Morning Setup
> I'm starting my shift and need to connect to M365 and refresh all tenant data. What's the best way to prepare the tool?

### End of Day Review
> Before I finish my shift, what should I check in the Activity Log to ensure all user creations were successful?

### Handover Documentation
> I need to document what user accounts I created today for the next shift. What information should I include?

### Weekly Maintenance
> What routine maintenance should I perform on the tool weekly to keep it running smoothly?

## Integration Prompts

### HR System Integration
> Our HR system exports new employee data. Can you help me format this data for bulk import into the M365 tool?

### Ticketing System
> I have a service desk ticket requesting a new user. Can you help me use the tool to fulfill this request efficiently?

### Documentation Updates
> After creating users, I need to update our internal documentation. What user details should I record?

## 🧪 Testing & Validation Prompts

### Testing Switch Tenant Functionality
> I want to test the Switch Tenant feature before using it with live clients. Can you help me run the testing scripts in the Tests folder?

### Debugging Tenant Data Issues
> I'm experiencing issues with tenant data not clearing properly. Can you help me use the debug scripts to troubleshoot this?

### Validating Button Visibility
> The Switch Tenant button isn't appearing in my GUI. Can you help me run the button testing scripts to diagnose the issue?

### Module Cache Problems
> I'm having issues with PowerShell modules not updating. Can you show me how to use the module reload scripts?

## Training Prompts

### New Team Member Training
> I'm training a new team member on this tool. Can you provide a step-by-step walkthrough of the basic user creation process?

### Switch Tenant Training
> I need to train my team on the new Switch Tenant functionality. Can you create a training outline covering the key steps and best practices?

### Best Practices
> What are the best practices for using this tool in a first-line support environment? What should I always check before creating users?

### Common Mistakes
> What are the most common mistakes first-line support makes when using this tool, and how can I avoid them?

## Emergency Prompts

### Urgent User Creation
> I have an urgent request to create a user account for someone starting today. What's the fastest way to get them set up?

### Tool Not Working
> The tool completely stopped working and I have users waiting. What basic troubleshooting steps should I try immediately?

### Bulk Import Failed
> My bulk import of 50 users failed halfway through. How do I figure out which users were created and which need to be retried?

---

## How to Use These Prompts

1. **Copy and paste** any relevant prompt into your conversation with Claude Code
2. **Customize the details** in brackets [like this] with your specific information
3. **Combine prompts** for complex scenarios (e.g., "Help me create a user AND add them to groups")
4. **Ask follow-up questions** if you need more detailed explanations
5. **Request step-by-step guidance** for any process you're unfamiliar with

Remember: Claude Code can see your entire project structure and files, so it can provide specific guidance based on your actual M365 User Provisioning Tool setup!
# Supabase Edge Function Setup

This guide will help you deploy the CORS proxy to Supabase Edge Functions.

## Prerequisites

1. **Supabase Account**: Sign up at https://supabase.com
2. **Supabase CLI**: Install via:
   ```bash
   npm install -g supabase
   ```

## Step-by-Step Deployment

### 1. Login to Supabase

```bash
supabase login
```

This will open a browser window for authentication.

### 2. Link to Your Project

Create a new project at https://app.supabase.com, then link it:

```bash
cd /Users/parthtrivedi/Documents/SE\ Toolkit/excel-data-addin
supabase link --project-ref YOUR_PROJECT_REF
```

Find your project ref in the Supabase dashboard URL: `https://app.supabase.com/project/YOUR_PROJECT_REF`

### 3. Deploy the Edge Function

```bash
supabase functions deploy fraud-proxy
```

### 4. Get Your Function URL

After deployment, your function URL will be:
```
https://YOUR_PROJECT_REF.supabase.co/functions/v1/fraud-proxy
```

### 5. Configure the Add-in

1. Open the Excel add-in
2. Go to the **Fraud Check** tab
3. Paste your function URL in the **Supabase Edge Function URL** field
4. The URL will be saved automatically

## Testing Locally (Optional)

To test the edge function locally before deploying:

```bash
supabase start
supabase functions serve fraud-proxy
```

Your local function will be available at:
```
http://localhost:54321/functions/v1/fraud-proxy
```

## Troubleshooting

### Function not deploying?
- Make sure you're in the correct directory
- Check that `supabase/functions/fraud-proxy/index.ts` exists
- Run `supabase functions list` to see deployed functions

### CORS errors persisting?
- The edge function has CORS enabled by default
- Check browser console for detailed error messages
- Verify the proxy URL is correct in the add-in settings

### API calls timing out?
- The timeout is set to 2 minutes
- Check Supabase function logs: `supabase functions logs fraud-proxy`
- Verify your Airia API key is valid

## Cost

Supabase Edge Functions are free for:
- 500,000 function invocations per month
- 2 million function execution seconds per month

Your fraud detection use case should easily fit within the free tier.

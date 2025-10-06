// Supabase Edge Function - CORS Proxy for Airia Fraud Detection API
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type, x-api-key, x-environment',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

serve(async (req) => {
  // Handle CORS preflight requests
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders });
  }

  try {
    // Parse request body
    const { environment, apiKey, data } = await req.json();

    if (!apiKey) {
      return new Response(
        JSON.stringify({ error: 'API key is required' }),
        { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      );
    }

    // Build API URL based on environment
    let apiUrl: string;

    if (environment === 'production') {
      apiUrl = 'https://api.airia.ai/v2/PipelineExecution/18546128-a4a6-411b-8b5c-23b64beaee01';
    } else if (environment === 'dev') {
      apiUrl = 'https://dev.api.airiadev.ai/v2/PipelineExecution/18546128-a4a6-411b-8b5c-23b64beaee01';
    } else {
      // Custom environment
      apiUrl = `https://${environment}.api.airia.ai/v2/PipelineExecution/18546128-a4a6-411b-8b5c-23b64beaee01`;
    }

    console.log(`Proxying request to: ${apiUrl}`);

    // Forward request to Airia API
    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: {
        'X-API-KEY': apiKey,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        userInput: JSON.stringify(data),
        asyncOutput: false,
      }),
    });

    // Get response data
    const responseData = await response.text();

    // Return response with CORS headers
    return new Response(responseData, {
      status: response.status,
      headers: {
        ...corsHeaders,
        'Content-Type': 'application/json',
      },
    });

  } catch (error) {
    console.error('Proxy error:', error);

    return new Response(
      JSON.stringify({
        error: 'Proxy error',
        message: error.message
      }),
      {
        status: 500,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      }
    );
  }
});
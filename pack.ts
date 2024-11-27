import * as coda from "@codahq/packs-sdk";

export const pack = coda.newPack();
pack.addNetworkDomain("coda.io");

pack.setUserAuthentication({
  type: coda.AuthenticationType.CodaApiHeaderBearerToken,
  shouldAutoAuthSetup: false,
  defaultConnectionRequirement: coda.ConnectionRequirement.Required,
  getConnectionName: async (context) => {
    // The connection is for maker+doc, so figure out
    let meData = await context.fetcher.fetch({
      method: 'GET',
      url: 'https://coda.io/apis/v1/whoami'
    });

    if (!context.endpoint) {
      return `${meData.body.name} (${meData.body.loginId}) - no doc selected`;
    }

    let docData = await context.fetcher.fetch({
      method: 'GET',
      url: context.endpoint
    });
    
    return `${meData.body.name} (${meData.body.loginId}) - ${docData.body.name}`;
  },
  postSetup: [{
    type: coda.PostSetupType.SetEndpoint,
    name: "SelectDoc",
    description: "Select the doc to connect this pack to",
    getOptions: async function (context) {
      if (!!context.invocationLocation.docId) {
        // Doc ID is still not removed from the SDK
        let docId = context.invocationLocation.docId;

        // Find the doc
        let docData = await context.fetcher.fetch({
          method: 'GET',
          url: `https://coda.io/apis/v1/docs/${docId}`
        });

        // Get me
        let meData = await context.fetcher.fetch({
          method: 'GET',
          url: 'https://coda.io/apis/v1/whoami'
        });

        // If I'm not the owner of the doc, crash
        if (docData.body.owner !== meData.body.loginId) {
          throw new coda.UserVisibleError("You are not the owner of this doc. Please ask the owner to install this pack on this doc");
        }

        return [{
          display: docData.body.name,
          value: docData.body.href
        }];
      }

      // Otherwise we don't have the doc ID in the context anymore, so just list all docs this user is owner of and collect until we have more pages
      let docItems : coda.MetadataFormulaObjectResultType[] = [];
      let continuationHref : string;
      do {
        let response = await context.fetcher.fetch({
          method: "GET",
          url: continuationHref || "https://coda.io/apis/v1/docs?isOwner=true&limit=100",
          cacheTtlSecs: 0
        });
        continuationHref = response.body.nextPageLink;

        for (let doc of response.body.items) {
          docItems.push({
            display: doc.name,
            value: doc.href
          });
        }

        console.log(continuationHref);
      } while (!!continuationHref);

      return docItems;
    }
  }]
});

const PAGE_SEARCH_FN = async (context, search, parameters) => {
  let pagesData = [];
  let continuationHref : string;

  do {
    let response = await context.fetcher.fetch({
      method: "GET",
      url: continuationHref || `${context.endpoint}/pages?limit=100`,
      cacheTtlSecs: 0
    });
    continuationHref = response.body.nextPageLink;
    
    for (let page of response.body.items) {
      pagesData.push(page);
    }
  } while (!!continuationHref)

  return coda.autocompleteSearchObjects(search, pagesData, "name", "id");
};

const ICON_SEARCH_FN = async (context, search, parameters) => {
  let iconsData: coda.MetadataFormulaObjectResultType[] = [];

  let response = await context.fetcher.fetch({
    method: "GET",
    url: coda.withQueryParams('https://coda.io/api/icons', {
      term: search,
      limit: 50
    })
  });
  
  for (let icon of response.body.icons) {
    iconsData.push({
      value: icon.name,
      display: icon.label
    });
  }

  return iconsData;
};

pack.addFormula({
  name: "ListPages",
  description: "Returns a list of all pages in the doc as [ID, name] pairs",
  parameters: [
    coda.makeParameter({
      type: coda.ParameterType.Number,
      name: "limit",
      description: "Maximum number of pages to return (default: 100)",
      optional: true
    })
  ],
  resultType: coda.ValueType.Array,
  items: coda.makeSchema({
    type: coda.ValueType.Array,
    items: { type: coda.ValueType.String }
  }),
  execute: async function ([limit], context) {
    const response = await context.fetcher.fetch({
      method: "GET",
      url: coda.withQueryParams(`${context.endpoint}/pages`, {
        limit: limit || 100
      }),
      cacheTtlSecs: 0
    });
    
    let pages = response.body.items.map(page => [page.id, page.name]);
    return pages;
  },
  cacheTtlSecs: 0
});

pack.addFormula({
  name: "AddPage",
  description: "Add a new page. Upon execution, returns an ID of the created page that you can then use to update it.",
  parameters: [
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "name",
      description: "Name of the page",
      optional: true
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "parent",
      description: "Parent of this new page",
      optional: true,
      autocomplete: PAGE_SEARCH_FN
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "subtitle",
      description: "Subtitle of the page",
      optional: true
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "iconName",
      description: "Name of the icon for this new page",
      optional: true,
      autocomplete: ICON_SEARCH_FN
    }),
    coda.makeParameter({
      type: coda.ParameterType.Image,
      name: "coverImage",
      description: "Cover image to use",
      optional: true
    }),
  ],
  resultType: coda.ValueType.String,
  isAction: true,
  execute: async function ([name, parentPageId, subtitle, iconName, imageUrl], context) {
    let payload = { name, parentPageId, subtitle, iconName, imageUrl };
    Object.keys(payload).forEach((k) => {
      if (payload[k] === undefined) {
        delete payload[k];
      }
    });
    let result = await context.fetcher.fetch({
      method: 'POST',
      url: `${context.endpoint}/pages`,
      body: JSON.stringify(payload),
      headers: {
        'Content-Type': 'application/json'
      }
    });
    return result.body.id
  },
});

pack.addFormula({
  name: "RenamePage",
  description: "Rename an existing page",
  parameters: [    
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "pageIdOrName",
      description: "ID or name of the page to rename. Prefer using IDs because names can change and there can be multiple pages with the same name. Use autocomplete to select the page, or use the result of the earlier AddPage() call, or switch on Developer Mode to get page IDs from the context menu",
      autocomplete: PAGE_SEARCH_FN
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "name",
      description: "New name of the page",
      optional: true
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "subtitle",
      description: "New subtitle of the page",
      optional: true
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "iconName",
      description: "New name of the icon for this page",
      optional: true,
      autocomplete: ICON_SEARCH_FN
    }),
    coda.makeParameter({
      type: coda.ParameterType.Image,
      name: "coverImage",
      description: "New cover image to use",
      optional: true
    }),
  ],
  resultType: coda.ValueType.String,
  isAction: true,
  execute: async function ([pageIdOrName, name, subtitle, iconName, imageUrl], context) {
    let payload = { name, subtitle, iconName, imageUrl };
    Object.keys(payload).forEach((k) => {
      if (payload[k] === undefined) {
        delete payload[k];
      }
    });
    let result = await context.fetcher.fetch({
      method: 'PUT',
      url: `${context.endpoint}/pages/${encodeURIComponent(pageIdOrName)}`,
      body: JSON.stringify(payload),
      headers: {
        'Content-Type': 'application/json'
      }
    });
    return result.body.id
  },
});
pack.addFormula({
  name: "CopyPage",
  description:
    "Create a copy of an existing page, including its content and properties",
  parameters: [
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "sourcePageIdOrName",
      description:
        "ID or name of the page to copy. Prefer using IDs because names can change and there can be multiple pages with the same name.",
      autocomplete: PAGE_SEARCH_FN,
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "newName",
      description: "Name for the copied page",
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "parentPageId",
      description: "Optional parent page for the copy",
      optional: true,
      autocomplete: PAGE_SEARCH_FN,
    }),
  ],
  resultType: coda.ValueType.String,
  isAction: true,
  execute: async function (
    [sourcePageIdOrName, newName, parentPageId],
    context
  ) {
    // Step 1: Get source page details
    const sourcePage = await context.fetcher.fetch({
      method: "GET",
      url: `${context.endpoint}/pages/${encodeURIComponent(sourcePageIdOrName)}`,
    });

    // Step 2: Begin content export of source page
    const exportResponse = await context.fetcher.fetch({
      method: "POST",
      url: `${context.endpoint}/pages/${encodeURIComponent(sourcePageIdOrName)}/export`,
      body: JSON.stringify({ outputFormat: "html" }),
      headers: { "Content-Type": "application/json" },
    });
    console.log("Export response:", exportResponse.body);

    // Step 3: Poll until export is complete
    const requestId = exportResponse.body.id;
    let downloadLink;
    while (true) {
      const statusResponse = await context.fetcher.fetch({
        method: "GET",
        url: `${context.endpoint}/pages/${encodeURIComponent(sourcePageIdOrName)}/export/${requestId}`,
      });
      console.log("Poll status response:", statusResponse.body);

      if (statusResponse.body.status === "complete") {
        downloadLink = statusResponse.body.downloadLink;
        console.log("Download link:", downloadLink);
        break;
      }

      // Wait before polling again
      await new Promise((resolve) => setTimeout(resolve, 1000));
    }

    // Step 4: Get the exported content
    const contentResponse = await context.fetcher.fetch({
      method: "GET",
      url: downloadLink,
    });
    const content = contentResponse.body;

    // Step 5: Create new page with content and properties from source
    const createPayload = {
      name: newName,
      parentPageId: parentPageId,
      subtitle: sourcePage.body.subtitle,
      iconName: sourcePage.body.icon?.name,
      imageUrl: sourcePage.body.image?.browserLink,
      pageContent: {
        type: "canvas",
        canvasContent: {
          format: "markdown",
          content: content,
        },
      },
    };

    // Remove undefined properties
    Object.keys(createPayload).forEach((k) => {
      if (createPayload[k] === undefined) {
        delete createPayload[k];
      }
    });

    const result = await context.fetcher.fetch({
      method: "POST",
      url: `${context.endpoint}/pages`,
      body: JSON.stringify(createPayload),
      headers: {
        "Content-Type": "application/json",
      },
    });

    return result.body.id;
  },
});

# SharePoint Sample Search Configurations
SharePoint Sample Search Configurations

   - <i>MapCrawledPropertyToManagedProperty.ps1</i>
        - <i>SearchMappingTemplate.xml</i> template to map cp to mp
        - <i>SearchMappingReset.xml</i> template to reset a cp mapping
   - <i>TagBoost-Max-SearchConfiguration.xml</i> - Boost equal to Title
   - <i>TagBoost-Medium-SearchConfiguration.xml</i> - Boost equal to SocialTag
   - <i>TagBoost-Low-SearchConfiguration.xml</i> - Boost equal to Body
   - <i>CrawlTimeSchema.xml</i> - Add CrawlTime managed property to see when an item was last crawled (only on-prem)
   - <i>HideFromDelveSchema.xml</i> - By using a Yes/No column named HideFromDelve in SharePoint you can omit them from the Delve board

## Description
Samples on how to map crawled properties to managed properties, and how to set the context level boost for a managed property in order to impact the BM25F score in search.

Any column which is of type managed metadata in SharePoint will get it's
values boosted with this search schema update.

For a detailed description of when and how to use these files check out [How to: Boost metadata in SharePoint search results]

[How to: Boost metadata in SharePoint search results]:http://techmikael.blogspot.com/2015/01/how-to-boost-metadata-in-sharepoint.html.

### Technical implementation
The search configuration file is creating a new managed property named <b>TagBoost</b>, mapping the catch-all crawled property <b>ows_taxId_MetadataAllTagsInfo</b> to it.

The different files are then associating the context level of <b>TabBoost</b> to the weight group of already existing managed properties in the default rank profile.

### Installation SharePoint Online
The files can be imported as search configurations files. For SharePoint Online do this at the tentant level from https://tenant-admin.sharepoint.com/_layouts/15/searchadmin/importsearchconfiguration.aspx?level=tenant

You can also import config files using the `Set-PnPSearchConfiguration` cmdlet.

### Installation SharePoint 2013 - On-premises

If you install at a site collection or in a multi tenancy environment follow the same procedure as for SharePoint Online. Remember that theany schema update done at a site or site collection level will only be valid for searches executed on that site collection. Meaning if content is residing on one site collection and search center on another, it won't work. Then you have to deploy the changes globally to the SSA.

If you want to do deploy the boost at the SSA level, you will manually have to create the managed property, change the weight group and map the crawled property as shown in [How to: Boost metadata in SharePoint search results].

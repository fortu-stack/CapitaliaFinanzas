import { defineCollection } from 'astro:content';
import { glob } from 'astro/loaders';
import { z } from 'astro/zod';

const blog = defineCollection({
	// Load Markdown and MDX files in the `src/content/blog/` directory.
	loader: glob({ base: './src/content/blog', pattern: '**/*.{md,mdx}' }),
	// Type-check frontmatter using a schema
	schema: ({ image }) =>
		z.object({
			title: z.string(),
			description: z.string(),
			// Transform string to Date object
			pubDate: z.coerce.date(),
			updatedDate: z.coerce.date().optional(),
			heroImage: z.optional(image()),
			// Nuevos campos para categorizar y dividir en Artículos, Calculadoras y FAQs
			category: z.enum(['vivienda', 'fiscalidad', 'inversion']).default('vivienda'),
			type: z.enum(['article', 'calculator', 'faq']).default('article'),
		}),
});

export const collections = { blog };

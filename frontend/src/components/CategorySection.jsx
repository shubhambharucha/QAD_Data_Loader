import { motion } from "framer-motion";

export default function CategorySection({ category, index }) {
  const IconComponent = category.icon;

  return (
    <section className="py-32 px-4 sm:px-6 lg:px-8 max-w-7xl mx-auto">
      {/* Section Header */}
      <motion.div
        initial={{ opacity: 0, x: index % 2 === 0 ? -30 : 30 }}
        whileInView={{ opacity: 1, x: 0 }}
        transition={{ duration: 0.8 }}
        viewport={{ once: true }}
        className="mb-16"
      >
        <div className="flex items-center gap-4 mb-4">
          <motion.div
            whileHover={{ rotate: 360 }}
            transition={{ duration: 0.6 }}
            className={`p-3 rounded-lg bg-gradient-to-br ${category.color}`}
          >
            <IconComponent size={28} className="text-white" />
          </motion.div>
          <div>
            <h2 className={`text-5xl font-black bg-gradient-to-r ${category.color} bg-clip-text text-transparent`}>
              {category.name}
            </h2>
            <p className="text-gray-400 text-lg">{category.description}</p>
          </div>
        </div>
      </motion.div>

      {/* Awards List */}
      <div className="space-y-4">
        {category.awards.map((award, awardIdx) => (
          <motion.div
            key={awardIdx}
            initial={{ opacity: 0, x: -20 }}
            whileInView={{ opacity: 1, x: 0 }}
            transition={{
              delay: awardIdx * 0.1,
              duration: 0.5,
            }}
            viewport={{ once: true }}
            whileHover={{ x: 8 }}
            className="group relative"
          >
            {/* Background glow on hover */}
            <div className={`absolute inset-0 bg-gradient-to-r ${category.color} rounded-xl opacity-0 group-hover:opacity-20 blur-xl transition-opacity duration-500`} />

            {/* Card */}
            <div className="relative bg-gradient-to-r from-slate-800 to-slate-900 border border-slate-700 group-hover:border-slate-600 rounded-xl p-6 transition-all duration-300">
              <div className="flex items-start justify-between">
                <div className="flex-1">
                  {/* Year Badge */}
                  <motion.span
                    initial={{ scale: 0 }}
                    whileInView={{ scale: 1 }}
                    transition={{ delay: 0.2 + awardIdx * 0.1 }}
                    className={`inline-block px-3 py-1 rounded-full text-xs font-bold text-white mb-3`}
                    style={{
                      background: `linear-gradient(135deg, var(--color-1) 0%, var(--color-2) 100%)`,
                      "--color-1": `rgb(${category.color.includes("blue") ? "59, 130, 246" : category.color.includes("pink") ? "236, 72, 153" : category.color.includes("purple") ? "147, 51, 234" : "249, 115, 22"})`,
                        "--color-2": `rgb(${category.color.includes("blue") ? "34, 211, 238" : category.color.includes("pink") ? "244, 63, 94" : category.color.includes("purple") ? "168, 85, 247" : "249, 115, 22"})`,
                    }}
                  >
                    {award.year}
                  </motion.span>

                  {/* Title */}
                  <h3 className="text-2xl font-black text-white mb-2 group-hover:text-transparent group-hover:bg-gradient-to-r group-hover:from-blue-400 group-hover:to-purple-400 group-hover:bg-clip-text transition-all duration-300">
                    {award.title}
                  </h3>

                  {/* Author */}
                  <p className="text-gray-400 flex items-center gap-2">
                    <span className="inline-block w-2 h-2 rounded-full" style={{background: category.color}} />
                    by {award.author}
                  </p>
                </div>

                {/* Vote count */}
                <motion.div
                  initial={{ scale: 0 }}
                  whileInView={{ scale: 1 }}
                  transition={{ delay: 0.3 + awardIdx * 0.1 }}
                  className="text-right ml-6"
                >
                  <p className={`text-3xl font-black bg-gradient-to-r ${category.color} bg-clip-text text-transparent`}>
                    {award.votes}
                  </p>
                  <p className="text-gray-500 text-xs font-semibold">Votes</p>
                </motion.div>
              </div>

              {/* Progress bar */}
              <motion.div
                initial={{ scaleX: 0 }}
                whileInView={{ scaleX: 1 }}
                transition={{
                  delay: 0.4 + awardIdx * 0.1,
                  duration: 0.6,
                }}
                className="mt-4 h-1 bg-gradient-to-r from-transparent via-slate-600 to-transparent rounded-full origin-left"
                style={{
                  background: `linear-gradient(90deg, transparent, var(--color-1), transparent)`,
                  "--color-1": category.color,
                }}
              />
            </div>
          </motion.div>
        ))}
      </div>
    </section>
  );
}
